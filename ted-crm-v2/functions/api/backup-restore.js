// Restauration d'une table depuis un backup KV.
// POST { date: 'YYYY-MM-DD', table: 'clients', confirm: 'RESTAURER' }
// - JWT Supabase requis (utilisateur CRM connecté)
// - Restauration par UPSERT (merge sur clé primaire) : les lignes du backup
//   sont recréées/écrasées, les lignes créées depuis ne sont PAS supprimées.
//   C'est volontairement non destructif.
import { fetchT, getSupa, supaHeaders, verifyUser, BACKUP_TABLES, secureJson } from '../_utils.js';

const BATCH = 200;

export async function onRequestPost(context) {
  const { request, env } = context;
  const user = await verifyUser(env, request);
  if (!user) return secureJson({ error: 'Non autorisé' }, { status: 401 });
  if (!env.BACKUPS) return secureJson({ error: 'Binding KV "BACKUPS" absent — voir MAINTENANCE.md' }, { status: 500 });

  let body;
  try { body = await request.json(); } catch { return secureJson({ error: 'JSON invalide' }, { status: 400 }); }
  const { date, table, confirm } = body || {};

  if (confirm !== 'RESTAURER') return secureJson({ error: 'Confirmation requise : confirm doit valoir "RESTAURER"' }, { status: 400 });
  if (!BACKUP_TABLES.includes(table)) return secureJson({ error: `Table inconnue. Tables restaurables : ${BACKUP_TABLES.join(', ')}` }, { status: 400 });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) return secureJson({ error: 'date invalide (YYYY-MM-DD)' }, { status: 400 });

  const { url, key, isService } = getSupa(env);
  if (!isService) return secureJson({ error: 'SUPABASE_SERVICE_ROLE_KEY manquante' }, { status: 500 });

  const raw = await env.BACKUPS.get(`backup:${date}`);
  if (!raw) return secureJson({ error: `Aucun backup pour le ${date}` }, { status: 404 });

  let backup;
  try { backup = JSON.parse(raw); } catch { return secureJson({ error: 'Backup corrompu (JSON illisible)' }, { status: 500 }); }
  const rows = backup.tables && backup.tables[table];
  if (!Array.isArray(rows)) return secureJson({ error: `La table ${table} est absente de ce backup` }, { status: 404 });

  let restored = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i += BATCH) {
    const batch = rows.slice(i, i + BATCH);
    try {
      const res = await fetchT(`${url}/rest/v1/${table}`, {
        method: 'POST',
        headers: supaHeaders(key, { Prefer: 'resolution=merge-duplicates,return=minimal' }),
        body: JSON.stringify(batch),
      }, 20000);
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${(await res.text()).slice(0, 200)}`);
      restored += batch.length;
    } catch (e) {
      errors.push(`lot ${i / BATCH + 1}: ${e.message}`);
    }
  }

  return secureJson({
    ok: errors.length === 0,
    table,
    date,
    restored,
    total: rows.length,
    errors: errors.length ? errors : undefined,
    restored_by: user.email || user.id,
  });
}
