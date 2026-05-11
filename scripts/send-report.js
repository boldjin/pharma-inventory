const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

// ── 날짜 유틸 ──────────────────────────────────────────────────────────────
function todayKST() {
  return new Date().toLocaleDateString('ko-KR', {
    timeZone: 'Asia/Seoul', year: 'numeric', month: '2-digit', day: '2-digit'
  }).replace(/\. /g, '-').replace('.', '');
}

// ── 현재 재고 계산 ──────────────────────────────────────────────────────────
function calcStock(products, transactions) {
  const stockMap = {};
  for (const p of products) stockMap[p.code] = 0;
  for (const t of transactions) {
    if (!stockMap.hasOwnProperty(t.product_code)) continue;
    if (t.type === '매입') stockMap[t.product_code] += t.qty;
    else if (t.type === '매출' || t.type === '폐기') stockMap[t.product_code] -= t.qty;
  }
  return stockMap;
}

// ── 스타일 헬퍼 ────────────────────────────────────────────────────────────
function headerStyle() {
  return { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1D4ED8' } }, alignment: { horizontal: 'center' }, border: { bottom: { style: 'thin' } } };
}

function applyHeaderStyle(ws, headers, row) {
  headers.forEach((_, ci) => {
    const addr = XLSX.utils.encode_cell({ r: row, c: ci });
    if (!ws[addr]) return;
    ws[addr].s = headerStyle();
  });
}

function setColWidths(ws, widths) {
  ws['!cols'] = widths.map(w => ({ wch: w }));
}

// ── 메인 ───────────────────────────────────────────────────────────────────
async function main() {
  const { SUPABASE_URL, SUPABASE_SERVICE_KEY, SUPABASE_USER_ID, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, RECIPIENT } = process.env;

  if (!SUPABASE_URL || !SUPABASE_SERVICE_KEY || !SUPABASE_USER_ID) {
    throw new Error('Supabase 환경변수가 설정되지 않았습니다. GitHub Secrets를 확인하세요.');
  }

  // 1. DB 조회
  const sb = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);
  const [{ data: products, error: e1 }, { data: transactions, error: e2 }] = await Promise.all([
    sb.from('products').select('*').eq('user_id', SUPABASE_USER_ID).order('code'),
    sb.from('transactions').select('*').eq('user_id', SUPABASE_USER_ID).order('date', { ascending: false }),
  ]);
  if (e1) throw new Error('품목 조회 실패: ' + e1.message);
  if (e2) throw new Error('거래 조회 실패: ' + e2.message);

  const stockMap = calcStock(products, transactions);
  const today = todayKST();
  const wb = XLSX.utils.book_new();

  // 2. 시트1: 현재재고 현황
  const stockHeaders = ['품목코드', '품목명', '규격', '단위', '카테고리', '제조사', '공급업체', '보관위치', '현재재고', '안전재고', '최대재고', '재고상태'];
  const stockRows = products.map(p => {
    const qty = stockMap[p.code] ?? 0;
    const status = qty <= 0 ? '재고없음' : qty <= p.safety_stock ? '부족' : qty >= p.max_stock ? '초과' : '정상';
    return [p.code, p.name, p.spec, p.unit, p.category, p.manufacturer, p.supplier, p.location, qty, p.safety_stock, p.max_stock, status];
  });
  const ws1 = XLSX.utils.aoa_to_sheet([stockHeaders, ...stockRows]);
  applyHeaderStyle(ws1, stockHeaders, 0);
  setColWidths(ws1, [12, 20, 14, 6, 14, 14, 14, 12, 10, 10, 10, 10]);
  XLSX.utils.book_append_sheet(wb, ws1, '현재재고');

  // 3. 시트2: 입출고기록
  const txHeaders = ['날짜', '유형', '품목코드', '품목명', 'LOT번호', '유효기간', '수량', '단가', '금액', '부서', '담당자', '거래처', '비고'];
  const txRows = transactions.map(t => [t.date, t.type, t.product_code, t.product_name, t.lot, t.expiry, t.qty, t.unit_price, t.amount, t.dept, t.manager, t.partner, t.note]);
  const ws2 = XLSX.utils.aoa_to_sheet([txHeaders, ...txRows]);
  applyHeaderStyle(ws2, txHeaders, 0);
  setColWidths(ws2, [12, 8, 12, 20, 14, 12, 8, 10, 12, 10, 10, 16, 20]);
  XLSX.utils.book_append_sheet(wb, ws2, '입출고기록');

  // 4. 시트3: 품목목록
  const prodHeaders = ['품목코드', '품목명', '규격', '단위', '카테고리', '제조사', '공급업체', '보관위치', '안전재고', '최대재고', '유효기간관리', '로트관리', '보험여부', '비고'];
  const prodRows = products.map(p => [p.code, p.name, p.spec, p.unit, p.category, p.manufacturer, p.supplier, p.location, p.safety_stock, p.max_stock, p.expiry_mgmt, p.lot_mgmt, p.insurance, p.note]);
  const ws3 = XLSX.utils.aoa_to_sheet([prodHeaders, ...prodRows]);
  applyHeaderStyle(ws3, prodHeaders, 0);
  setColWidths(ws3, [12, 20, 14, 6, 14, 14, 14, 12, 10, 10, 12, 10, 10, 20]);
  XLSX.utils.book_append_sheet(wb, ws3, '품목목록');

  const xlsxBuf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  // 5. 이메일 발송
  const transporter = nodemailer.createTransport({
    host: SMTP_HOST || 'smtp.naver.com',
    port: parseInt(SMTP_PORT || '587'),
    secure: parseInt(SMTP_PORT || '587') === 465,
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });

  await transporter.sendMail({
    from: `"의약품 재고관리" <${SMTP_USER}>`,
    to: RECIPIENT || 'forzaackj@naver.com',
    subject: `[재고관리] ${today} 일일 재고 리포트`,
    html: `
      <div style="font-family:sans-serif; max-width:500px;">
        <h2 style="color:#1d4ed8;">📊 ${today} 재고 현황</h2>
        <table style="border-collapse:collapse; width:100%;">
          <tr style="background:#dbeafe;">
            <td style="padding:8px 12px; font-weight:bold;">전체 품목</td>
            <td style="padding:8px 12px;">${products.length}개</td>
          </tr>
          <tr>
            <td style="padding:8px 12px; font-weight:bold;">재고 부족</td>
            <td style="padding:8px 12px; color:#dc2626;">${products.filter(p => (stockMap[p.code] ?? 0) <= p.safety_stock && p.safety_stock > 0).length}개</td>
          </tr>
          <tr style="background:#dbeafe;">
            <td style="padding:8px 12px; font-weight:bold;">금일 입출고</td>
            <td style="padding:8px 12px;">${transactions.filter(t => t.date === today).length}건</td>
          </tr>
        </table>
        <p style="margin-top:16px; color:#6b7280; font-size:13px;">상세 내용은 첨부 엑셀 파일을 확인해 주세요.</p>
      </div>
    `,
    attachments: [{
      filename: `재고관리_${today}.xlsx`,
      content: xlsxBuf,
    }],
  });

  console.log(`✅ 리포트 발송 완료: ${today}`);
}

main().catch(err => {
  console.error('❌ 발송 실패:', err.message);
  process.exit(1);
});
