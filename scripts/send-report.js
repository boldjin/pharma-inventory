const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

// ── 날짜 유틸 ──────────────────────────────────────────────────────────────
function todayKST() {
  return new Date().toLocaleDateString('ko-KR', {
    timeZone: 'Asia/Seoul', year: 'numeric', month: '2-digit', day: '2-digit'
  }).replace(/\. /g, '-').replace('.', '');
}
function daysUntil(dateStr) {
  if (!dateStr) return null;
  const target = new Date(dateStr);
  const now = new Date();
  return Math.floor((target - now) / (1000 * 60 * 60 * 24));
}

// ── 한국 공휴일 (대체공휴일 포함) ─────────────────────────────────────────
// 매년 1월에 갱신 필요 — 음력 공휴일과 대체공휴일은 해마다 날짜가 바뀜
const KR_HOLIDAYS = new Set([
  // 2026
  '2026-01-01', // 신정
  '2026-02-16', '2026-02-17', '2026-02-18', // 설날 연휴
  '2026-03-01', '2026-03-02', // 삼일절 + 대체
  '2026-05-05', // 어린이날
  '2026-05-24', '2026-05-25', // 부처님오신날 + 대체
  '2026-06-06', // 현충일
  '2026-08-15', '2026-08-17', // 광복절 + 대체
  '2026-09-24', '2026-09-25', '2026-09-26', '2026-09-28', // 추석 연휴 + 대체
  '2026-10-03', '2026-10-05', // 개천절 + 대체
  '2026-10-09', // 한글날
  '2026-12-25', // 크리스마스
  // 2027
  '2027-01-01',
  '2027-02-06', '2027-02-07', '2027-02-08', '2027-02-09', // 설날 + 대체
  '2027-03-01',
  '2027-05-05',
  '2027-05-13',
  '2027-06-07', // 현충일 대체 (6/6 일요일)
  '2027-08-16', // 광복절 대체 (8/15 일요일)
  '2027-09-14', '2027-09-15', '2027-09-16',
  '2027-10-04', // 개천절 대체
  '2027-10-11', // 한글날 대체
  '2027-12-27', // 크리스마스 대체
]);

function todayKSTDate() {
  // YYYY-MM-DD 형태 (KST 기준)
  const fmt = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Seoul', year: 'numeric', month: '2-digit', day: '2-digit' });
  return fmt.format(new Date());
}
function todayKSTDayOfWeek() {
  // KST 기준 요일: 0=일, 6=토
  const parts = new Intl.DateTimeFormat('en-US', { timeZone: 'Asia/Seoul', weekday: 'short' }).format(new Date());
  return { Sun:0, Mon:1, Tue:2, Wed:3, Thu:4, Fri:5, Sat:6 }[parts];
}
function shouldSkipToday() {
  const dow = todayKSTDayOfWeek();
  if (dow === 0 || dow === 6) return { skip: true, reason: '주말' };
  if (KR_HOLIDAYS.has(todayKSTDate())) return { skip: true, reason: '공휴일' };
  return { skip: false };
}

// ── 집계 ──────────────────────────────────────────────────────────────────
function calcStock(products, transactions) {
  const stockMap = {};
  const lastInMap = {};
  const lastOutMap = {};
  const stockValueMap = {};
  // 평균 매입단가 계산용
  const purchaseAgg = {};

  for (const p of products) {
    stockMap[p.code] = 0;
    stockValueMap[p.code] = 0;
    purchaseAgg[p.code] = { qty: 0, amount: 0 };
  }
  // 오래된 순으로 정렬해서 집계 일관성 확보
  const sorted = [...transactions].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  for (const t of sorted) {
    if (!stockMap.hasOwnProperty(t.product_code)) continue;
    if (t.type === '매입') {
      stockMap[t.product_code] += t.qty;
      purchaseAgg[t.product_code].qty += t.qty;
      purchaseAgg[t.product_code].amount += (t.amount || 0);
      if (!lastInMap[t.product_code] || t.date > lastInMap[t.product_code]) lastInMap[t.product_code] = t.date;
    } else if (t.type === '매출' || t.type === '폐기') {
      stockMap[t.product_code] -= t.qty;
      if (!lastOutMap[t.product_code] || t.date > lastOutMap[t.product_code]) lastOutMap[t.product_code] = t.date;
    }
  }
  // 재고금액 = 평균 매입단가 * 현재재고
  for (const code of Object.keys(stockMap)) {
    const agg = purchaseAgg[code];
    const avg = agg.qty > 0 ? agg.amount / agg.qty : 0;
    stockValueMap[code] = Math.round(avg * Math.max(0, stockMap[code]));
  }
  return { stockMap, lastInMap, lastOutMap, stockValueMap };
}

function stockStatus(qty, safety, max) {
  if (qty <= 0) return '⚠부족';
  if (safety > 0 && qty <= safety) return '⚠부족';
  if (max > 0 && qty >= max) return '●초과';
  if (safety > 0 && qty <= safety * 1.5) return '◎주의';
  return '●정상';
}

// ── 스타일 / 시트 헬퍼 ──────────────────────────────────────────────────────
const TITLE_STYLE  = { font: { bold: true, sz: 14, color: { rgb: '1D4ED8' } }, alignment: { horizontal: 'left', vertical: 'center' } };
const HEADER_STYLE = { font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 }, fill: { fgColor: { rgb: '1D4ED8' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: { top:{style:'thin',color:{rgb:'CBD5E1'}}, bottom:{style:'thin',color:{rgb:'CBD5E1'}}, left:{style:'thin',color:{rgb:'CBD5E1'}}, right:{style:'thin',color:{rgb:'CBD5E1'}} } };
const CELL_BORDER  = { border: { top:{style:'thin',color:{rgb:'E5E7EB'}}, bottom:{style:'thin',color:{rgb:'E5E7EB'}}, left:{style:'thin',color:{rgb:'E5E7EB'}}, right:{style:'thin',color:{rgb:'E5E7EB'}} } };

function applyStyles(ws, title, headers, dataRowCount) {
  // 타이틀(0행)
  const titleAddr = XLSX.utils.encode_cell({ r: 0, c: 0 });
  if (ws[titleAddr]) ws[titleAddr].s = TITLE_STYLE;
  // 헤더(2행 = index 2 because row1 is blank-ish; we'll use row index 1 in this script)
  headers.forEach((_, ci) => {
    const addr = XLSX.utils.encode_cell({ r: 1, c: ci });
    if (ws[addr]) ws[addr].s = HEADER_STYLE;
  });
  // 데이터 행
  for (let r = 2; r < 2 + dataRowCount; r++) {
    for (let c = 0; c < headers.length; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (ws[addr]) ws[addr].s = { ...CELL_BORDER, alignment: { vertical: 'center' } };
    }
  }
  // 타이틀 셀 머지
  ws['!merges'] = ws['!merges'] || [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } });
  ws['!rows'] = ws['!rows'] || [];
  ws['!rows'][0] = { hpt: 24 };
  ws['!rows'][1] = { hpt: 20 };
}

function buildSheet(title, headers, rows, widths) {
  // 1행: 타이틀, 2행: 헤더, 3행~: 데이터
  const aoa = [[title], headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = widths.map(w => ({ wch: w }));
  applyStyles(ws, title, headers, rows.length);
  return ws;
}

// ── 메인 ───────────────────────────────────────────────────────────────────
async function main() {
  const { SUPABASE_URL, SUPABASE_SERVICE_KEY, SUPABASE_USER_ID, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, RECIPIENT, FORCE_SEND } = process.env;
  if (!SUPABASE_URL || !SUPABASE_SERVICE_KEY || !SUPABASE_USER_ID) {
    throw new Error('Supabase 환경변수가 설정되지 않았습니다.');
  }

  // 주말/공휴일 발송 제외 (FORCE_SEND=1 이면 우회)
  if (FORCE_SEND !== '1') {
    const { skip, reason } = shouldSkipToday();
    if (skip) {
      console.log(`⏭️  ${todayKSTDate()} (${reason}) — 발송 생략`);
      return;
    }
  }

  // 1. DB 조회
  const sb = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);
  const [{ data: products, error: e1 }, { data: transactions, error: e2 }] = await Promise.all([
    sb.from('products').select('*').eq('user_id', SUPABASE_USER_ID).order('code'),
    sb.from('transactions').select('*').eq('user_id', SUPABASE_USER_ID).order('date', { ascending: false }),
  ]);
  if (e1) throw new Error('품목 조회 실패: ' + e1.message);
  if (e2) throw new Error('거래 조회 실패: ' + e2.message);

  const { stockMap, lastInMap, lastOutMap, stockValueMap } = calcStock(products, transactions);
  const today = todayKST();
  const wb = XLSX.utils.book_new();

  // ── 시트1: 기준정보 (Master Data) ────────────────────────────────────────
  const masterHeaders = ['품목코드', '품목명', '규격', '단위', '분류', '제조사', '공급처', '보관위치', '안전재고', '최대재고', '유효기간관리', 'LOT관리', '보험/비보험', '비고'];
  const masterRows = products.map(p => [
    p.code, p.name, p.spec || '', p.unit || '', p.category || '', p.manufacturer || '', p.supplier || '', p.location || '',
    p.safety_stock || 0, p.max_stock || 0, p.expiry_mgmt || 'N', p.lot_mgmt || 'N', p.insurance || '보험', p.note || ''
  ]);
  const ws1 = buildSheet('■ 의료용 소모품 기준정보 (Master Data)', masterHeaders, masterRows,
    [12, 22, 14, 6, 14, 14, 14, 12, 10, 10, 12, 10, 12, 16]);
  XLSX.utils.book_append_sheet(wb, ws1, '기준정보');

  // ── 시트2: 입출고기록 (Transaction Log) ──────────────────────────────────
  const txHeaders = ['일자', '구분', '품목코드', '품목명', 'LOT번호', '유효기간', '수량', '단가', '금액', '부서', '담당자', '거래처/사유', '비고'];
  const txRows = transactions.map(t => [
    t.date || '', t.type || '', t.product_code || '', t.product_name || '', t.lot || '', t.expiry || '',
    t.qty || 0, t.unit_price || 0, t.amount || 0, t.dept || '', t.manager || '', t.partner || '', t.note || ''
  ]);
  const ws2 = buildSheet('■ 입출고 기록 (Transaction Log)', txHeaders, txRows,
    [12, 8, 12, 22, 14, 12, 8, 10, 12, 12, 10, 20, 20]);
  XLSX.utils.book_append_sheet(wb, ws2, '입출고기록');

  // ── 시트3: 현재재고 (Current Stock Status) ───────────────────────────────
  const stockHeaders = ['품목코드', '품목명', '단위', '안전재고', '현재재고', '재고상태', '재고금액', '최근입고일', '최근사용일', '비고'];
  const stockRows = products.map(p => {
    const qty = stockMap[p.code] ?? 0;
    return [
      p.code, p.name, p.unit || '', p.safety_stock || 0, qty,
      stockStatus(qty, p.safety_stock || 0, p.max_stock || 0),
      stockValueMap[p.code] || 0,
      lastInMap[p.code] || '',
      lastOutMap[p.code] || '',
      p.note || ''
    ];
  });
  const ws3 = buildSheet('■ 현재재고 현황 (Current Stock Status)', stockHeaders, stockRows,
    [12, 22, 6, 10, 10, 10, 14, 14, 14, 16]);
  XLSX.utils.book_append_sheet(wb, ws3, '현재재고');

  // ── 시트4: 유효기간관리 (Expiry Date Management) ─────────────────────────
  const expiryHeaders = ['품목코드', '품목명', 'LOT번호', '유효기간', '입고수량', '사용/폐기수량', '잔량', '잔여일수', '상태'];
  // LOT별 집계
  const lotMap = {};
  for (const t of transactions) {
    if (!t.lot || !t.expiry) continue;
    const key = `${t.product_code}__${t.lot}`;
    if (!lotMap[key]) {
      lotMap[key] = {
        code: t.product_code, name: t.product_name, lot: t.lot, expiry: t.expiry,
        inQty: 0, outQty: 0
      };
    }
    if (t.type === '매입') lotMap[key].inQty += (t.qty || 0);
    else if (t.type === '매출' || t.type === '폐기') lotMap[key].outQty += (t.qty || 0);
  }
  const expiryRows = Object.values(lotMap).map(l => {
    const remain = l.inQty - l.outQty;
    const days = daysUntil(l.expiry);
    let status = '●정상';
    if (days === null) status = '';
    else if (days < 0) status = '■만료';
    else if (days <= 30) status = '⚠임박';
    else if (days <= 90) status = '◎주의';
    return [l.code, l.name, l.lot, l.expiry, l.inQty, l.outQty, remain, days, status];
  });
  const ws4 = buildSheet('■ 유효기간 관리 (Expiry Date Management)', expiryHeaders,
    expiryRows.length > 0 ? expiryRows : [['', '', '', '', '', '', '', '', '(LOT 관리 품목 거래 없음)']],
    [12, 22, 14, 12, 10, 12, 8, 10, 10]);
  XLSX.utils.book_append_sheet(wb, ws4, '유효기간관리');

  // ── 시트5: 사용안내 ──────────────────────────────────────────────────────
  const guideAoa = [
    ['📋 의료용 소모품 재고관리대장 사용안내'],
    [],
    ['■ 시트 구성'],
    ['', '기준정보',    '품목 마스터 데이터. 신규 품목 추가 시 이 시트에 먼저 등록하세요.'],
    ['', '입출고기록',  '모든 입고/사용/폐기/반품/조정 내역을 일자별로 기록합니다.'],
    ['', '현재재고',    '품목별 현재 재고 수량, 재고금액, 안전재고 상태를 자동 집계합니다.'],
    ['', '유효기간관리', 'LOT별 유효기간과 잔여일수, 만료 상태를 관리합니다.'],
    ['', '사용안내',    '본 가이드 시트입니다.'],
    [],
    ['■ 재고상태 표기'],
    ['', '⚠부족',  '현재재고 ≤ 안전재고 또는 재고없음'],
    ['', '◎주의',  '안전재고 < 현재재고 ≤ 안전재고 × 1.5'],
    ['', '●정상',  '적정 범위'],
    ['', '●초과',  '현재재고 ≥ 최대재고'],
    [],
    ['■ 유효기간 상태'],
    ['', '■만료', '유효기간 경과'],
    ['', '⚠임박', '잔여일수 ≤ 30일'],
    ['', '◎주의', '잔여일수 ≤ 90일'],
    ['', '●정상', '잔여일수 > 90일'],
    [],
    ['■ 자동 발송'],
    ['', '매일 오전 9시', `${process.env.RECIPIENT || ''} 외 등록된 이메일로 자동 발송됩니다.`],
  ];
  const ws5 = XLSX.utils.aoa_to_sheet(guideAoa);
  ws5['!cols'] = [{ wch: 4 }, { wch: 16 }, { wch: 60 }];
  if (ws5['A1']) ws5['A1'].s = { font: { bold: true, sz: 14, color: { rgb: '1D4ED8' } } };
  ['A3','A10','A16','A22'].forEach(addr => { if (ws5[addr]) ws5[addr].s = { font: { bold: true, color: { rgb: '1D4ED8' } } }; });
  XLSX.utils.book_append_sheet(wb, ws5, '사용안내');

  const xlsxBuf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  // ── 이메일 발송 ──────────────────────────────────────────────────────────
  const recipients = (RECIPIENT || 'forzaackj@naver.com')
    .split(',').map(s => s.trim()).filter(Boolean);

  const transporter = nodemailer.createTransport({
    host: SMTP_HOST || 'smtp.naver.com',
    port: parseInt(SMTP_PORT || '587'),
    secure: parseInt(SMTP_PORT || '587') === 465,
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });

  const lowStockCount = products.filter(p => {
    const q = stockMap[p.code] ?? 0;
    return (p.safety_stock || 0) > 0 && q <= p.safety_stock;
  }).length;
  const todayTxCount = transactions.filter(t => t.date === today).length;

  await transporter.sendMail({
    from: `"의약품 재고관리" <${SMTP_USER}>`,
    to: recipients,
    subject: `[재고관리] ${today} 일일 재고 리포트`,
    html: `
      <div style="font-family:-apple-system,BlinkMacSystemFont,'Malgun Gothic',sans-serif; max-width:560px;">
        <h2 style="color:#1d4ed8; margin-bottom:8px;">📊 ${today} 재고 현황</h2>
        <table style="border-collapse:collapse; width:100%; font-size:14px;">
          <tr style="background:#dbeafe;"><td style="padding:10px 14px; font-weight:600; width:160px;">전체 품목</td><td style="padding:10px 14px;">${products.length}개</td></tr>
          <tr><td style="padding:10px 14px; font-weight:600; border-top:1px solid #e5e7eb;">재고 부족</td><td style="padding:10px 14px; color:#dc2626; border-top:1px solid #e5e7eb;">${lowStockCount}개</td></tr>
          <tr style="background:#dbeafe;"><td style="padding:10px 14px; font-weight:600;">금일 입출고</td><td style="padding:10px 14px;">${todayTxCount}건</td></tr>
          <tr><td style="padding:10px 14px; font-weight:600; border-top:1px solid #e5e7eb;">전체 거래기록</td><td style="padding:10px 14px; border-top:1px solid #e5e7eb;">${transactions.length}건</td></tr>
        </table>
        <p style="margin-top:18px; color:#6b7280; font-size:13px; line-height:1.6;">
          첨부된 엑셀에는 <b>기준정보 · 입출고기록 · 현재재고 · 유효기간관리 · 사용안내</b> 5개 시트가 포함되어 있습니다.<br>
          웹 시스템: <a href="https://boldjin.github.io/pharma-inventory">https://boldjin.github.io/pharma-inventory</a>
        </p>
      </div>
    `,
    attachments: [{
      filename: `재고관리_${today}.xlsx`,
      content: xlsxBuf,
    }],
  });

  console.log(`✅ 리포트 발송 완료: ${today} → ${recipients.join(', ')}`);
}

main().catch(err => {
  console.error('❌ 발송 실패:', err.message);
  process.exit(1);
});
