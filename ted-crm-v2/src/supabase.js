import { createClient } from '@supabase/supabase-js';
import { createResilientFetch, logError } from './lib/db';

const supabaseUrl = process.env.REACT_APP_SUPABASE_URL;
const supabaseKey = process.env.REACT_APP_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseKey) {
  throw new Error('Variables Supabase manquantes (REACT_APP_SUPABASE_URL / REACT_APP_SUPABASE_ANON_KEY). Vérifiez votre fichier .env ou les variables de build Cloudflare.');
}

// Client unique de l'app : tous les appels passent par le fetch résilient
// (timeout 10 s, retry exponentiel, log error_logs en échec final).
export const supabase = createClient(supabaseUrl, supabaseKey, {
  global: {
    fetch: createResilientFetch({
      onFinalError: (e, url) =>
        logError((e && e.message) || 'fetch failed', `supabase:${String(url).split('?')[0].slice(-80)}`),
    }),
  },
});
