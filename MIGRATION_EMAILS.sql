CREATE TABLE IF NOT EXISTS emails_envoyes (
  id uuid default gen_random_uuid() primary key,
  objet text not null,
  message text,
  nb_destinataires int default 0,
  destinataires jsonb,
  envoye_par text,
  created_at timestamptz default now(),
  statut text default 'envoye'
);
ALTER TABLE emails_envoyes ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth all emails" ON emails_envoyes FOR ALL USING (auth.role() = 'authenticated');
