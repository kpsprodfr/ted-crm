// Sauvegarde quotidienne des tables Supabase vers Cloudflare KV.
//
// - GET  (pg_cron, 02:00 UTC) : idempotent — si le backup du jour existe déjà,
//   no-op silencieux (pas d'email). Aucun secret nécessaire : rejouer l'appel
//   ne fait rien, la réponse ne contient que des compteurs.
// - POST (bouton « Backup maintenant » du CRM, JWT Supabase requis) :
//   force la re-création du backup du jour.
//
// Stockage : KV `BACKUPS` (namespace ted-crm-backups), clé backup:YYYY-MM-DD,
// rétention 30 jours, checksum SHA-256, email récap à com.astegal@gmail.com.
import { fetchT, getSupa, supaHeaders, sendBrevoEmail, sha256Hex, verifyUser, BACKUP_TABLES } from '../_utils.js';

const RETENTION = 30;

export async function onRequest(context) {
  const { request, env } = context;
  if (request.method !== 'GET' && request.method !== 'POST') {
    return Response.json({ error: 'Méthode non autorisée' }, { status: 405 });
  }

  if (!env.BACKUPS) {
    return Response.json({ ok: false, error: 'Binding KV "BACKUPS" absent — à ajouter dans Cloudflare Pages → Settings → Bindings (namespace ted-crm-backups). Voir MAINTENANCE.md.' }, { status: 500 });
  }

  const { url, key, isService } = getSupa(env);
  if (!isService) {
    return Response.json({ ok: false, error: 'SUPABASE_SERVICE_ROLE_KEY manquante — export impossible.' }, { status: 500 });
  }

  const today = new Date().toISOString().split('T')[0];
  const kvKey = `backup:${today}`;

  let force = false;
  if (request.method === 'POST') {
    const user = await verifyUser(env, request);
    if (!user) return Response.json({ error: 'Non autorisé' }, { status: 401 });
    force = true;
  }

  // Idempotence : le backup du jour existe déjà → no-op (sauf force)
  if (!force) {
    const existing = await env.BACKUPS.getWithMetadata(kvKey, { type: 'stream' });
    if (existing && existing.value) {
      try { await existing.value.cancel(); } catch { /* noop */ }
      return Response.json({ ok: true, skipped: true, date: today, info: 'Backup du jour déjà présent' });
    }
  }

  // Export de toutes les tables (clé service → RLS bypassée)
  const tables = {};
  const counts = {};
  const errors = [];
  for (const t of BACKUP_TABLES) {
    try {
      const res = await fetchT(`${url}/rest/v1/${t}?select=*&limit=50000`, { headers: supaHeaders(key) }, 20000);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const rows = await res.json();
      tables[t] = rows;
      counts[t] = Array.isArray(rows) ? rows.length : 0;
    } catch (e) {
      errors.push(`${t}: ${e.message}`);
      counts[t] = -1;
    }
  }

  const payload = { date: today, exported_at: new Date().toISOString(), tables, counts, errors };
  const json = JSON.stringify(payload);
  const checksum = await sha256Hex(json);
  const size = json.length;

  await env.BACKUPS.put(kvKey, json, {
    metadata: { date: today, exported_at: payload.exported_at, counts, size, checksum, errors: errors.length },
  });

  // Rétention : on garde les 30 derniers
  let deleted = 0;
  try {
    const list = await env.BACKUPS.list({ prefix: 'backup:' });
    const keys = list.keys.map((k) => k.name).sort(); // backup:YYYY-MM-DD → tri lexico = tri chrono
    while (keys.length - deleted > RETENTION) {
      await env.BACKUPS.delete(keys[deleted]);
      deleted += 1;
    }
  } catch (e) {
    errors.push(`rétention: ${e.message}`);
  }

  // Email récapitulatif
  const lignes = BACKUP_TABLES.map((t) => `<tr><td style="padding:4px 12px;border-bottom:1px solid #eee;">${t}</td><td style="padding:4px 12px;border-bottom:1px solid #eee;text-align:right;">${counts[t] >= 0 ? counts[t] + ' lignes' : '⚠️ erreur'}</td></tr>`).join('');
  const emailRes = await sendBrevoEmail(env, {
    to_email: 'com.astegal@gmail.com',
    to_name: 'Le TED',
    subject: `${errors.length ? '⚠️' : '✅'} Backup TED CRM ${today} — ${Object.values(counts).filter(n => n >= 0).reduce((a, b) => a + b, 0)} lignes`,
    html: `<h2 style="font-family:Arial,sans-serif;">Backup quotidien TED CRM</h2>
<p style="font-family:Arial,sans-serif;">Date : <strong>${today}</strong> · Taille : <strong>${(size / 1024).toFixed(1)} Ko</strong> · Checksum SHA-256 : <code style="font-size:11px;">${checksum.slice(0, 16)}…</code></p>
<table style="font-family:Arial,sans-serif;font-size:13px;border-collapse:collapse;">${lignes}</table>
${errors.length ? `<p style="color:#c00;font-family:Arial,sans-serif;">Erreurs : ${errors.join(' · ')}</p>` : ''}
<p style="font-family:Arial,sans-serif;color:#888;font-size:12px;">Stocké dans Cloudflare KV (rétention 30 jours). Restauration : CRM → Système.</p>`,
  });

  return Response.json({
    ok: errors.length === 0,
    date: today,
    counts,
    size,
    checksum,
    deleted_old: deleted,
    email_sent: emailRes.ok,
    errors: errors.length ? errors : undefined,
  });
}
