-- ============================================================
-- 의약품 재고관리 시스템 - Supabase 초기 설정 SQL
-- Supabase 대시보드 > SQL Editor 에서 전체 실행
-- ============================================================

-- 1. Products 테이블
create table public.products (
  code          text primary key,
  name          text not null,
  spec          text default '',
  unit          text default '',
  category      text default '',
  manufacturer  text default '',
  supplier      text default '',
  location      text default '',
  safety_stock  integer default 0,
  max_stock     integer default 0,
  expiry_mgmt   text default 'N',
  lot_mgmt      text default 'N',
  insurance     text default '보험',
  note          text default '',
  user_id       uuid references auth.users not null,
  created_at    timestamptz default now()
);

-- 2. Transactions 테이블
create table public.transactions (
  id            bigserial primary key,
  date          text not null,
  type          text not null,
  product_code  text not null,
  product_name  text not null,
  lot           text default '',
  expiry        text default '',
  qty           integer not null default 0,
  unit_price    numeric not null default 0,
  amount        numeric not null default 0,
  dept          text default '',
  manager       text default '',
  partner       text default '',
  note          text default '',
  user_id       uuid references auth.users not null,
  created_at    timestamptz default now()
);

-- 3. Row Level Security 활성화
alter table public.products     enable row level security;
alter table public.transactions enable row level security;

-- 4. RLS 정책 (본인 데이터만 접근)
create policy "owner_products" on public.products
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);

create policy "owner_transactions" on public.transactions
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);

-- 5. Realtime 활성화
alter publication supabase_realtime add table public.products;
alter publication supabase_realtime add table public.transactions;

-- 6. 일일 메일 발송 로그 (중복 발송 방지)
create table if not exists public.daily_email_log (
  date         text primary key,        -- 'YYYY-MM-DD' (KST 기준)
  sent_at      timestamptz default now(),
  recipients   text default ''
);
-- service_role 키로만 접근 (RLS 없이도 anon key는 차단됨)
alter table public.daily_email_log enable row level security;
create policy "service_role_only" on public.daily_email_log
  for all using (false) with check (false);  -- anon/authenticated 모두 차단, service_role은 RLS 우회
