// Reprend les emails en attente dans email_queue (échecs Brevo).
// Déclenché toutes les heures par pg_cron côté Supabase (net.http_get).
//
// Endpoint volontairement SANS secret mais sûr par construction :
//  - claim atomique en DB (pending → processing) : un rejeu concurrent ne peut
//    pas renvoyer deux fois le même email (le 2e appel ne "claim" rien) ;
//  - la réponse ne contient que des compteurs, jamais de données ;
//  - nécessite SUPABASE_SERVICE_ROLE_KEY côté Cloudflare pour lire la file
//    (RLS : anon = écriture seule) — sinon no-op explicite.
import { fetchT, sendBrevoEmail, supaHeaders, getSupa } from '../_utils.js';

const MAX_ATTEMPTS = 5;
const BATCH = 20;
const STALE_PROCESSING_MIN = 15;

export async function onRequest(context) {
  const { env } = context;
  const { url, key, isService } = getSupa(env);
  if (!isService) {
    return Response.json({
      processed: 0,
      error: 'SUPABASE_SERVICE_ROLE_KEY manquante dans Cloudflare — impossible de lire la file (RLS). Voir MAINTENANCE.md.',
    });
  }

  const nowIso = new Date().toISOString();

  // 0. Récupération des items bloqués en 'processing' (crash d'un run précédent)
  try {
    const staleBefore = new Date(Date.now() - STALE_PROCESSING_MIN * 60000).toISOString();
    await fetchT(`${url}/rest/v1/email_queue?status=eq.processing&claimed_at=lt.${staleBefore}`, {
      method: 'PATCH',
      headers: supaHeaders(key, { Prefer: 'return=minimal' }),
      body: JSON.stringify({ status: 'pending', claimed_at: null }),
    });
  } catch { /* non bloquant */ }

  // 1. Sélection des candidats
  let candidates = [];
  try {
    const res = await fetchT(
      `${url}/rest/v1/email_queue?status=eq.pending&attempts=lt.${MAX_ATTEMPTS}&order=created_at.asc&limit=${BATCH}&select=id`,
      { headers: supaHeaders(key) }
    );
    if (!res.ok) throw new Error(`Supabase ${res.status}`);
    candidates = await res.json();
  } catch (e) {
    return Response.json({ processed: 0, error: `Lecture file impossible: ${e.message}` }, { status: 502 });
  }
  if (!candidates.length) return Response.json({ ok: true, pending: 0, sent: 0, retried: 0, failed_permanently: 0 });

  // 2. Claim atomique : seules les lignes ENCORE pending sont flippées et retournées.
  //    Un appel concurrent obtient [] et s'arrête → aucun double envoi possible.
  let claimed = [];
  try {
    const ids = candidates.map((c) => c.id).join(',');
    const res = await fetchT(
      `${url}/rest/v1/email_queue?id=in.(${ids})&status=eq.pending`,
      {
        method: 'PATCH',
        headers: supaHeaders(key, { Prefer: 'return=representation' }),
        body: JSON.stringify({ status: 'processing', claimed_at: nowIso }),
      }
    );
    if (!res.ok) throw new Error(`Supabase ${res.status}`);
    claimed = await res.json();
  } catch (e) {
    return Response.json({ processed: 0, error: `Claim impossible: ${e.message}` }, { status: 502 });
  }

  const report = { pending: claimed.length, sent: 0, retried: 0, failed_permanently: 0 };

  // 3. Envoi + statut final
  for (const item of claimed) {
    const result = await sendBrevoEmail(env, {
      to_email: item.to_email,
      to_name: item.to_name,
      subject: item.subject,
      html: item.html,
    });

    const attempts = (item.attempts || 0) + 1;
    let patch;
    if (result.ok) {
      patch = { status: 'sent', attempts, last_attempt: nowIso, claimed_at: null, error_message: null };
      report.sent += 1;
    } else if (attempts >= MAX_ATTEMPTS) {
      patch = { status: 'failed', attempts, last_attempt: nowIso, claimed_at: null, error_message: `Abandon après ${attempts} tentatives — Brevo ${result.status}: ${result.detail}`.slice(0, 1000) };
      report.failed_permanently += 1;
    } else {
      patch = { status: 'pending', attempts, last_attempt: nowIso, claimed_at: null, error_message: `Brevo ${result.status}: ${result.detail}`.slice(0, 1000) };
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
