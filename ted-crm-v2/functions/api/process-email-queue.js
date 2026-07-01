// Reprend les emails en attente dans email_queue (échecs Brevo).
// Appelé toutes les heures par le workflow GitHub Actions .github/workflows/cron.yml
// (Cloudflare Pages n'exécute pas les crons wrangler — voir MAINTENANCE.md).
// Protégé : Authorization: Bearer <BACKUP_SECRET>.
// Nécessite SUPABASE_SERVICE_ROLE_KEY pour lire la file (RLS : anon = écriture seule).
import { fetchT, sendBrevoEmail, supaHeaders, getSupa } from '../_utils.js';

const MAX_ATTEMPTS = 5;
const BATCH = 20;

export async function onRequest(context) {
  const { request, env } = context;
  if (request.method !== 'GET' && request.method !== 'POST') {
    return Response.json({ error: 'Méthode non autorisée' }, { status: 405 });
  }

  const token = (request.headers.get('Authorization') || '').replace('Bearer ', '');
  if (!env.BACKUP_SECRET || token !== env.BACKUP_SECRET) {
    return Response.json({ error: 'Non autorisé' }, { status: 401 });
  }

  const { url, key, isService } = getSupa(env);
  if (!isService) {
    return Response.json({
      processed: 0,
      error: 'SUPABASE_SERVICE_ROLE_KEY manquante dans Cloudflare — impossible de lire la file (RLS). Voir MAINTENANCE.md.',
    });
  }

  let queue = [];
  try {
    const res = await fetchT(
      `${url}/rest/v1/email_queue?status=eq.pending&attempts=lt.${MAX_ATTEMPTS}&order=created_at.asc&limit=${BATCH}`,
      { headers: supaHeaders(key) }
    );
    if (!res.ok) throw new Error(`Supabase ${res.status}`);
    queue = await res.json();
  } catch (e) {
    return Response.json({ processed: 0, error: `Lecture file impossible: ${e.message}` }, { status: 502 });
  }

  const report = { pending: queue.length, sent: 0, retried: 0, failed_permanently: 0 };

  for (const item of queue) {
    const result = await sendBrevoEmail(env, {
      to_email: item.to_email,
      to_name: item.to_name,
      subject: item.subject,
      html: item.html,
    });

    const attempts = (item.attempts || 0) + 1;
    let patch;
    if (result.ok) {
      patch = { status: 'sent', attempts, last_attempt: new Date().toISOString(), error_message: null };
      report.sent += 1;
    } else if (attempts >= MAX_ATTEMPTS) {
      patch = { status: 'failed', attempts, last_attempt: new Date().toISOString(), error_message: `Abandon après ${attempts} tentatives — Brevo ${result.status}: ${result.detail}`.slice(0, 1000) };
      report.failed_permanently += 1;
    } else {
      patch = { status: 'pending', attempts, last_attempt: new Date().toISOString(), error_message: `Brevo ${result.status}: ${result.detail}`.slice(0, 1000) };
      report.retried += 1;
    }

    try {
      await fetchT(`${url}/rest/v1/email_queue?id=eq.${item.id}`, {
        method: 'PATCH',
        headers: supaHeaders(key, { Prefer: 'return=minimal' }),
        body: JSON.stringify(patch),
      });
    } catch (e) {
      console.error('[process-email-queue] update raté pour', item.id, e.message);
    }
  }

  return Response.json({ ok: true, ...report });
}
