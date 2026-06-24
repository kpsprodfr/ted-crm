create table if not exists reservations (
  id uuid default gen_random_uuid() primary key,
  client_id uuid references clients(id) on delete set null,
  date date not null,
  service text not null check (service in ('midi','soir')),
  heure text,
  nb_personnes int not null check (nb_personnes >= 1),
  occasion text,
  commentaire_client text,
  note_interne text,
  statut text not null default 'attente' check (statut in ('attente','rappeler','confirmee','refusee','annulee','venue','absente')),
  raison_refus text,
  raison_annulation text,
  source text default 'manuel',
  date_rappel timestamptz,
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  traited_at timestamptz,
  traited_by text
);
alter table clients add column if not exists tel_normalise text;
alter table clients add column if not exists email_normalise text;
alter table clients add column if not exists source text default 'manuel';
alter table clients add column if not exists derniere_visite date;
alter table clients add column if not exists nb_reservations int default 0;
alter table clients add column if not exists nb_venues int default 0;
alter table clients add column if not exists nb_absences int default 0;
alter table clients add column if not exists nb_annulations int default 0;
create table if not exists parametres (
  id int primary key default 1,
  nom_etablissement text default 'Le TED',
  email_notification text,
  email_expediteur text,
  capacite_midi int default 300,
  capacite_soir int default 300,
  jours_ouverts text default 'lun,mar,mer,jeu,ven,sam',
  emails_actifs boolean default false,
  msg_confirmation text,
  msg_refus text,
  msg_accuse text
);
insert into parametres (id) values (1) on conflict (id) do nothing;
create index if not exists idx_resa_client on reservations(client_id);
create index if not exists idx_resa_date on reservations(date);
create index if not exists idx_resa_statut on reservations(statut);
create index if not exists idx_clients_tel on clients(tel_normalise);
alter table reservations enable row level security;
alter table parametres enable row level security;
create policy "auth read resa" on reservations for select using (auth.role() = 'authenticated');
create policy "auth insert resa" on reservations for insert with check (true);
create policy "auth update resa" on reservations for update using (auth.role() = 'authenticated');
create policy "auth params" on parametres for all using (auth.role() = 'authenticated');
