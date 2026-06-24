CREATE TABLE IF NOT EXISTS fcm_tokens (
  id uuid default gen_random_uuid() primary key,
  token text not null unique,
  user_id uuid references auth.users(id),
  created_at timestamptz default now()
);
ALTER TABLE fcm_tokens ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth fcm" ON fcm_tokens FOR ALL USING (auth.role() = 'authenticated');
