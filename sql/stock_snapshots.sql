-- 在処管理：時点別の在庫スナップショット
-- Supabase → SQL Editor に貼って実行してください（1回だけ）
-- https://supabase.com/dashboard/project/yukcdqnnevomhdcpsvms/sql

create table if not exists stock_snapshots (
  id            bigserial primary key,
  snapshot_date date        not null,   -- 「いつ時点」か。例: 2026-07-31
  label         text,                   -- 例: 2026年7月末 棚卸し
  kind          text        not null,   -- fabric / product / sakamoto
  fabric_no     text,                   -- 生地用
  color         text,                   -- 生地用
  code          text,                   -- 商品用
  name          text,                   -- 商品用 / 坂本繊維の品番
  qty           numeric     not null default 0,
  location_id   text,
  location_name text,                   -- 撮った時点の拠点名を固定保存（工場名が変わっても過去は読める）
  unit          text,                   -- 反 / 枚
  memo          text,
  created_at    timestamptz not null default now()
);

create index if not exists stock_snapshots_date_kind_idx
  on stock_snapshots (snapshot_date desc, kind);

-- 同じ日付を撮り直せるように、日付＋種別で消してから入れ直す運用にする
alter table stock_snapshots enable row level security;

drop policy if exists "stock_snapshots anon read"   on stock_snapshots;
drop policy if exists "stock_snapshots anon write"  on stock_snapshots;
drop policy if exists "stock_snapshots anon update" on stock_snapshots;
drop policy if exists "stock_snapshots anon delete" on stock_snapshots;

create policy "stock_snapshots anon read"   on stock_snapshots for select to anon using (true);
create policy "stock_snapshots anon write"  on stock_snapshots for insert to anon with check (true);
create policy "stock_snapshots anon update" on stock_snapshots for update to anon using (true) with check (true);
create policy "stock_snapshots anon delete" on stock_snapshots for delete to anon using (true);
