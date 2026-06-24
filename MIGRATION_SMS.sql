CREATE TABLE IF NOT EXISTS sms_envoyes (
  id uuid default gen_random_uuid() primary key,
  message text,
  nb_destinataires int default 0,
  destinataires jsonb,
  envoye_par text,
  created_at timestamptz default now()
);
ALTER TABLE sms_envoyes ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth all sms" ON sms_envoyes FOR ALL USING (auth.role() = 'authenticated');
