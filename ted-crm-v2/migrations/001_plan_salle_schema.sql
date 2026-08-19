-- Plan de salle — 1/2 : schéma
-- Les coordonnées sont en CENTIMÈTRES (unités réelles), pas en pixels ni en
-- pourcentage : le plan est indépendant de la taille de l'écran.
-- À exécuter dans Supabase → SQL Editor (projet mwpfaytccypvdrgapptk).

create or replace function public.plan_touch_updated_at()
returns trigger
language plpgsql
security invoker
set search_path = public, pg_temp
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

-- ── Espaces ────────────────────────────────────────────────────────────────
create table if not exists public.salles (
  id         uuid primary key default gen_random_uuid(),
  nom        text not null,
  type       text not null default 'salle'
             check (type in ('salle','etage','terrasse','bar')),
  ordre      integer not null default 0,
  largeur    integer not null default 1200,   -- cm
  hauteur    integer not null default 800,    -- cm
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists salles_ordre_idx on public.salles (ordre, created_at);

-- ── Tables ─────────────────────────────────────────────────────────────────
create table if not exists public.tables_salle (
  id           uuid primary key default gen_random_uuid(),
  salle_id     uuid not null references public.salles(id) on delete cascade,
  libelle      text not null,
  forme        text not null default 'ronde'
               check (forme in ('ronde','ovale','carree','rectangle','banquette')),
  x            integer not null default 0,
  y            integer not null default 0,
  largeur      integer not null default 70,
  hauteur      integer not null default 70,
  rotation     integer not null default 0,
  nb_couverts  integer not null default 2 check (nb_couverts between 1 and 30),
  statut       text not null default 'active'
               check (statut in ('active','hors_service')),
  notes        text,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now()
);
create index if not exists tables_salle_salle_idx on public.tables_salle (salle_id);

-- ── Décors (bar, entrée, escalier, cuisine…) ───────────────────────────────
-- Séparés des tables : pas de couverts, pas d'assignation, jamais comptés.
create table if not exists public.decors_salle (
  id         uuid primary key default gen_random_uuid(),
  salle_id   uuid not null references public.salles(id) on delete cascade,
  libelle    text not null,
  type       text not null default 'bloc'
             check (type in ('bar','entree','escalier','cuisine','wc','bloc')),
  x          integer not null default 0,
  y          integer not null default 0,
  largeur    integer not null default 100,
  hauteur    integer not null default 100,
  rotation   integer not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists decors_salle_salle_idx on public.decors_salle (salle_id);

-- ── Placement des réservations ─────────────────────────────────────────────
alter table public.reservations
  add column if not exists table_id uuid references public.tables_salle(id) on delete set null;
create index if not exists reservations_table_id_idx on public.reservations (table_id);

-- Toutes les tables d'une réservation (une seule en temps normal, plusieurs
-- quand on combine des tables pour un grand groupe).
create table if not exists public.reservations_tables (
  reservation_id uuid not null references public.reservations(id) on delete cascade,
  table_id       uuid not null references public.tables_salle(id) on delete cascade,
  created_at     timestamptz not null default now(),
  primary key (reservation_id, table_id)
);
create index if not exists reservations_tables_table_idx on public.reservations_tables (table_id);

-- reservations.table_id suit la table principale (la première assignée), pour
-- que le code simple n'ait qu'une colonne à lire.
create or replace function public.plan_sync_table_principale()
returns trigger
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  cible uuid := coalesce(new.reservation_id, old.reservation_id);
begin
  update public.reservations r
     set table_id = (
       select rt.table_id from public.reservations_tables rt
        where rt.reservation_id = cible
        order by rt.created_at, rt.table_id
        limit 1)
   where r.id = cible;
  return null;
end;
$$;

drop trigger if exists trg_plan_sync_table_principale on public.reservations_tables;
create trigger trg_plan_sync_table_principale
  after insert or delete on public.reservations_tables
  for each row execute function public.plan_sync_table_principale();

-- ── updated_at ─────────────────────────────────────────────────────────────
drop trigger if exists trg_salles_touch on public.salles;
create trigger trg_salles_touch before update on public.salles
  for each row execute function public.plan_touch_updated_at();

drop trigger if exists trg_tables_salle_touch on public.tables_salle;
create trigger trg_tables_salle_touch before update on public.tables_salle
  for each row execute function public.plan_touch_updated_at();

drop trigger if exists trg_decors_salle_touch on public.decors_salle;
create trigger trg_decors_salle_touch before update on public.decors_salle
  for each row execute function public.plan_touch_updated_at();

-- ── RLS : le plan est interne. anon n'a AUCUN accès. ───────────────────────
alter table public.salles              enable row level security;
alter table public.tables_salle        enable row level security;
alter table public.decors_salle        enable row level security;
alter table public.reservations_tables enable row level security;

drop policy if exists "auth all salles" on public.salles;
create policy "auth all salles" on public.salles
  to authenticated using (true) with check (true);

drop policy if exists "auth all tables_salle" on public.tables_salle;
create policy "auth all tables_salle" on public.tables_salle
  to authenticated using (true) with check (true);

drop policy if exists "auth all decors_salle" on public.decors_salle;
create policy "auth all decors_salle" on public.decors_salle
  to authenticated using (true) with check (true);

drop policy if exists "auth all reservations_tables" on public.reservations_tables;
create policy "auth all reservations_tables" on public.reservations_tables
  to authenticated using (true) with check (true);

revoke all on public.salles, public.tables_salle, public.decors_salle,
              public.reservations_tables from anon;
