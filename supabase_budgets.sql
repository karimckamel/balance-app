-- Run this in Supabase SQL Editor

create table budgets (
  id bigint primary key generated always as identity,
  user_id uuid references auth.users not null,
  category text not null,
  year int not null,
  month int, -- NULL = base budget (applies to all months unless overridden)
  amount numeric not null default 0,
  created_at timestamptz default now(),
  unique(user_id, category, year, month)
);

alter table budgets enable row level security;

create policy "users see own budgets" on budgets for all using (auth.uid() = user_id);
