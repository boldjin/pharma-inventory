# 의약품 재고관리 시스템

## 프로젝트 개요
브라우저 기반 의약품 재고관리 웹앱. 데이터는 Supabase PostgreSQL에 저장되며, 로그인 후 어느 기기에서든 동일한 데이터에 접근 가능.

## 기술 스택
- **순수 HTML/CSS/JS** — 빌드 도구 없음, 단일 파일 `index.html`
- **Chart.js 4.4.0** — CDN
- **Supabase JS SDK v2** — CDN, 인증 + PostgreSQL + Realtime

## 개발 서버 실행
```bash
python3 -m http.server 8080
# → http://localhost:8080
```

## Supabase 연동 설정
`index.html` 상단 SUPABASE CONFIG 섹션에 본인 프로젝트 값 입력:
```js
const SUPABASE_URL  = 'https://xxxx.supabase.co';
const SUPABASE_ANON = 'eyJ...';
```
초기 DB 스키마는 `supabase-setup.sql` 실행 (Supabase > SQL Editor).

## DB 스키마 (Supabase PostgreSQL)
| 테이블 | 주요 컬럼 |
|---|---|
| `products` | code(PK), name, safety_stock, max_stock, expiry_mgmt, lot_mgmt, user_id |
| `transactions` | id(bigserial PK), date, type, product_code, unit_price, amount, user_id |

- RLS 적용: 본인 `user_id` 데이터만 접근
- Realtime 활성화: INSERT/UPDATE/DELETE 즉시 화면 반영

## 아키텍처 패턴
- **In-memory cache**: `_products[]`, `_transactions[]` — 로그인 시 전체 로드, 이후 읽기는 동기(캐시)
- **쓰기**: DB 저장 → `loadAllData()` → 화면 갱신
- **Realtime**: Supabase `postgres_changes` 구독 → 자동 캐시 갱신 + 화면 갱신

## 주요 탭 구조
| 탭 | data-tab | 렌더 함수 |
|---|---|---|
| 대시보드 | dashboard | renderDashboard() |
| 현재재고 | stock | renderStock() |
| 매입현황 | purchase | renderPurchase() |
| 매출현황 | sales | renderSales() |
| 입출고기록 | records | renderRecords() |
| 품목관리 | products | renderProductsPage() |

## 코드 구조 (index.html 내 섹션)
- `SUPABASE CONFIG` — URL/KEY 상수
- `IN-MEMORY CACHE` — `_products`, `_transactions`, `_userId`
- `DB ↔ JS 변환` — `productToDb/FromDb`, `txToDb/FromDb` (snake_case ↔ camelCase)
- `SUPABASE DB 함수` — `loadAllData`, `dbUpsertProduct`, `dbDeleteProduct`, `dbInsertTransaction`, `dbUpdateTransaction`, `dbDeleteTransaction`
- `SEED DATA` — 최초 로그인 시 빈 DB에 자동 삽입
- `AUTH UI` — 로그인 오버레이, `doLogin()`, `doLogout()`
- `CRUD: TRANSACTION` — `saveTransaction()`, `deleteTransaction()` (async)
- `CRUD: PRODUCT` — `saveProduct()`, `deleteProduct()` (async)
- `REALTIME` — `subscribeRealtime()` (user_id 필터)
- `INIT` — `initApp(user)`, `sb.auth.onAuthStateChange()`

## 개발 규칙
- 빌드 없이 `index.html` 직접 수정
- localStorage 완전 제거 — 모든 데이터는 Supabase
- DB 쓰기 함수는 항상 async/await
- 새 데이터 타입 추가 시 `supabase-setup.sql`도 함께 업데이트
- 한국어 UI 유지
