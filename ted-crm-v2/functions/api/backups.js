// Liste les backups disponibles dans KV (page Système du CRM).
// JWT Supabase requis — les métadonnées seulement, jamais le contenu.
import { verifyUser, secureJson } from '../_utils.js';

export async function onRequestGet(context) {
  const { request, env } = context;
  const user = await verifyUser(env, request);
  if (!user) return secureJson({ error: 'Non autorisé' }, { status: 401 });
  if (!env.BACKUPS) {
    return secureJson({ backups: [], error: 'Binding KV "BACKUPS" absent — voir MAINTENANCE.md' });
  }
  const list = await env.BACKUPS.list({ prefix: 'backup:' });
  const backups = list.keys
    .map((k) => ({ key: k.name, date: k.name.replace('backup:', ''), ...(k.metadata || {}) }))
    .sort((a, b) => (a.date < b.date ? 1 : -1));
  return secureJson({ backups });
}
