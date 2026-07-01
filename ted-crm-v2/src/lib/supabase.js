// Doublon historique : on ré-exporte le client unique pour éviter deux
// connexions Supabase divergentes. Utiliser src/supabase.js partout.
export { supabase } from '../supabase';
