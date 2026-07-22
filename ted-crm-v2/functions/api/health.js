// État de santé du système — appelé par la pastille du CRM (60 s),
// et par pg_cron toutes les 15 min avec ?notify=1 (alerte email si down,
// dédupliquée à 1/heure via KV).
//
// Réponse publique : statuts et latences uniquement (jamais de données).
// Statuts : ok | degraded | down.
import { fetchT, getSupa, supaHeaders, sendBrevoEmail, queueEmail, secureJson } from '../_utils.js';

const CRITICAL_TABLES = ['clients', 'reservations', 'roue_gains', 'parametres'];

export async function onRequestGet(context) {
  const { request, env } = context;
  const t0 = Date.now();
  const components = {};

  // 1. Variables d'environnement
  const missingVars = [];
  if (!env.BREVO_API_KEY) missingVars.push('BREVO_API_KEY');
  if (!env.SUPABASE_SERVICE_ROLE_KEY && !env.SUPABASE_SERVICE_KEY) missingVars.push('SUPABASE_SERVICE_ROLE_KEY');
  components.env_vars = missingVars.length
    ? { status: 'down', detail: `Manquantes : ${missingVars.join(', ')}` }
    : { status: 'ok' };

  // 2. Supabase (latence + accès aux tables critiques)
  const { url, key } = getSupa(env);
  try {
    const t = Date.now();
    const checks = await Promise.all(
      CRITICAL_TABLES.map((tb) =>
        fetchT(`${url}/rest/v1/${tb}?select=id&limit=1`, { headers: supaHeaders(key) }, 8000)
          .then((r) => ({ table: tb, ok: r.ok, status: r.status }))
          .catch((e) => ({ table: tb, ok: false, error: e.message }))
      )
    );
    const latency = Date.now() - t;
    const ko = checks.filter((c) => !c.ok);
    components.supabase = ko.length
      ? { status: ko.length === CRITICAL_TABLES.length ? 'down' : 'degraded', latency_ms: latency, detail: `Tables KO : ${ko.map((c) => c.table).join(', ')}` }
      : { status: latency > 3000 ? 'degraded' : 'ok', latency_ms: latency, tables: CRITICAL_TABLES.length };
  } catch (e) {
    components.supabase = { status: 'down', detail: e.message };
  }

  // 3. Brevo (clé valide ?)
  if (env.BREVO_API_KEY) {
    try {
      const t = Date.now();
      const res = await fetchT('https://api.brevo.com/v3/account', { headers: { 'api-key': env.BREVO_API_KEY } }, 8000);
      components.brevo = res.ok
        ? { status: 'ok', latency_ms: Date.now() - t }
        : { status: 'down', detail: `HTTP ${res.status} — clé invalide ou compte bloqué` };
    } catch (e) {
      components.brevo = { status: 'down', detail: e.message };
    }
  } else {
    components.brevo = { status: 'down', detail: 'BREVO_API_KEY manquante (les emails partent en file d’attente)' };
  }

  // 4. Backups (binding KV + fraîcheur du dernier backup)
  // Lecture d'une clé unique (get, quota ~100k/j) plutôt qu'un list() (quota
  // ~1000/j) — cette route est appelée toutes les 60s par la pastille du CRM,
  // un list() ici épuisait le quota KV gratuit en une journée.
  if (!env.BACKUPS) {
    components.backups = { status: 'degraded', detail: 'Binding KV "BACKUPS" absent — backups inactifs (voir MAINTENANCE.md)' };
  } else {
    try {
      const raw = await env.BACKUPS.get('backups:meta:latest');
      const meta = raw ? JSON.parse(raw) : null;
      const last = meta ? meta.date : null;
      const ageDays = last ? Math.floor((Date.now() - new Date(last + 'T02:00:00Z').getTime()) / 86400000) : null;
      components.backups = !last
        ? { status: 'degraded', detail: 'Aucun backup encore présent' }
        : ageDays > 2
          ? { status: 'degraded', detail: `Dernier backup : ${last} (${ageDays} j)` }
          : { status: 'ok', last_backup: last, count: meta.count };
    } catch (e) {
      components.backups = { status: 'degraded', detail: e.message };
    }
  }

  // Statut global : down si Supabase down ou emails impossibles (Brevo + env KO)
  const statuses = Object.values(components).map((c) => c.status);
  let global = 'ok';
  if (statuses.includes('degraded')) global = 'degraded';
  if (components.supabase.status === 'down' || components.brevo.status === 'down' || components.env_vars.status === 'down') global = 'down';

  const body = { status: global, checked_at: new Date().toISOString(), duration_ms: Date.now() - t0, components };

  // Alerte email (cron 15 min avec ?notify=1) — max 1/heure via TTL KV
  const notify = new URL(request.url).searchParams.get('notify') === '1';
  if (notify && global === 'down') {
    let muted = false;
    if (env.BACKUPS) {
      muted = !!(await env.BACKUPS.get('health:last-alert'));
      if (!muted) await env.BACKUPS.put('health:last-alert', new Date().toISOString(), { expirationTtl: 3600 });
    }
    if (!muted) {
      const detail = Object.entries(components)
        .map(([name, c]) => `<tr><td style="padding:4px 12px;">${name}</td><td style="padding:4px 12px;font-weight:bold;color:${c.status === 'ok' ? '#15803d' : c.status === 'degraded' ? '#b45309' : '#b91c1c'};">${c.status}</td><td style="padding:4px 12px;color:#666;">${c.detail || ''}</td></tr>`)
        .join('');
      const email = {
        to_email: 'com.astegal@gmail.com',
        to_name: 'Le TED',
        subject: `🚨 ALERTE TED CRM — système ${global.toUpperCase()}`,
        html: `<h2 style="font-family:Arial,sans-serif;color:#b91c1c;">Un composant du TED CRM est en panne</h2>
<table style="font-family:Arial,sans-serif;font-size:13px;border-collapse:collapse;">${detail}</table>
<p style="font-family:Arial,sans-serif;color:#888;font-size:12px;">Vérifié le ${body.checked_at}. Prochaine alerte au plus tôt dans 1 h. Détail : CRM → Système.</p>`,
      };
      const sent = await sendBrevoEmail(env, email);
      if (!sent.ok) await queueEmail(env, { ...email, error_message: `alerte santé — Brevo ${sent.status}: ${sent.detail}` });
    }
  }

  return secureJson(body, { status: global === 'down' ? 503 : 200 });
}
