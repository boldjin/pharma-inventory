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

-- 6. 담당자 정산 테이블
create table if not exists public.settlements (
  id                bigserial primary key,
  date              text not null,             -- YYYY-MM-DD
  hospital_name     text not null default '',  -- 병원명
  product_code      text not null default '',
  product_name      text not null default '',
  qty               integer not null default 0,        -- 병원 납품수량
  unit_price        numeric not null default 0,        -- 병원납품가
  amount            numeric not null default 0,        -- 매출액
  per_unit_amount   numeric not null default 0,        -- 개당 정산금액
  margin_rate       numeric not null default 0,        -- 마진률 (0~1)
  manager           text not null default '',          -- 담당자
  settlement_amount numeric not null default 0,        -- 담당자 정산금액
  option_type       text default '프리랜서',           -- 프리랜서 / 사업자
  note              text default '',
  transaction_id    bigint references public.transactions(id) on delete set null,  -- 매출 거래 FK (1:1)
  user_id           uuid references auth.users not null,
  created_at        timestamptz default now()
);
create index if not exists idx_settlements_transaction on public.settlements(transaction_id);
alter table public.settlements enable row level security;
create policy "owner_settlements" on public.settlements
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);
alter publication supabase_realtime add table public.settlements;

-- 7. 일정 캘린더 테이블
create table if not exists public.schedules (
  id          bigserial primary key,
  date        text not null,              -- YYYY-MM-DD
  person      text not null default '',   -- 담당자명
  title       text not null default '',
  schedule_type text default '외근',       -- 외근/회의/기타
  start_time  text default '',             -- HH:MM
  end_time    text default '',
  location    text default '',
  note        text default '',
  shared      boolean default true,        -- 다른 사람과 공유 여부
  user_id     uuid references auth.users not null,
  created_at  timestamptz default now()
);
alter table public.schedules enable row level security;
create policy "owner_schedules" on public.schedules
  for all using (auth.uid() = user_id) with check (auth.uid() = user_id);
alter publication supabase_realtime add table public.schedules;

-- 8. 일일 메일 발송 로그 (중복 발송 방지)
create table if not exists public.daily_email_log (
  date         text primary key,        -- 'YYYY-MM-DD' (KST 기준)
  sent_at      timestamptz default now(),
  recipients   text default ''
);
-- service_role 키로만 접근 (RLS 없이도 anon key는 차단됨)
alter table public.daily_email_log enable row level security;
create policy "service_role_only" on public.daily_email_log
  for all using (false) with check (false);  -- anon/authenticated 모두 차단, service_role은 RLS 우회
