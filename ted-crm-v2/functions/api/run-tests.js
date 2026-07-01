// Suite de tests du système — rapport JSON succès/échec par test.
// Accès : header X-Test-Key: <BACKUP_SECRET>  OU  JWT CRM (bouton page Système).
// Tests : connexion Supabase, lecture des tables, écriture/suppression réelle
// (sur error_logs), clé Brevo, variables d'env, KV backups, santé globale.
import { fetchT, getSupa, supaHeaders, verifyUser, guard, secureJson, BACKUP_TABLES } from '../_utils.js';

export async function onRequest(context) {
  const { request, env } = context;
  if (request.method !== 'GET' && request.method !== 'POST') {
    return secureJson({ error: 'Méthode non autorisée' }, { status: 405 });
  }

  const blocked = await guard(env, request, { limit: 6, bucket: 'run-tests' });
  if (blocked) return blocked;

  // Auth : secret d'infra OU utilisateur CRM connecté
  const testKey = request.headers.get('X-Test-Key') || '';
  const bySecret = env.BACKUP_SECRET && testKey === env.BACKUP_SECRET;
  const byUser = bySecret ? null : await verifyUser(env, request);
  if (!bySecret && !byUser) return secureJson({ error: 'Non autorisé' }, { status: 401 });

  const t0 = Date.now();
  const tests = [];
  const add = (name, ok, detail = '', ms = null) => tests.push({ name, ok, detail, ms });
  const { url, key, isService } = getSupa(env);

  // 1. Variables d'environnement
  add('env: BREVO_API_KEY', !!env.BREVO_API_KEY, env.BREVO_API_KEY ? 'présente' : 'MANQUANTE');
  add('env: SUPABASE_SERVICE_ROLE_KEY', isService, isService ? 'présente' : 'MANQUANTE — file emails, backups et tests écriture inopérants');
  add('env: binding KV BACKUPS', !!env.BACKUPS, env.BACKUPS ? 'présent' : 'ABSENT — backups et rate limiting inactifs');

  // 2. Connexion Supabase + lecture de chaque table sauvegardée
  let t = Date.now();
  try {
    const results = await Promise.all(
      BACKUP_TABLES.map((tb) =>
        fetchT(`${url}/rest/v1/${tb}?select=*&limit=1`, { headers: supaHeaders(key) }, 8000)
          .then((r) => ({ tb, ok: r.ok, status: r.status }))
          .catch((e) => ({ tb, ok: false, detail: e.message }))
      )
    );
    const ko = results.filter((r) => !r.ok);
    add('supabase: connexion + lecture 12 tables', ko.length === 0, ko.length ? `KO : ${ko.map((r) => r.tb).join(', ')}` : `${results.length} tables lisibles`, Date.now() - t);
  } catch (e) {
    add('supabase: connexion + lecture 12 tables', false, e.message, Date.now() - t);
  }

  // 3. Écriture + relecture + suppression réelles (table error_logs, inoffensif)
  if (isService) {
    t = Date.now();
    try {
      const marker = `run-tests-${Date.now()}`;
      const ins = await fetchT(`${url}/rest/v1/error_logs`, {
        method: 'POST',
        headers: supaHeaders(key, { Prefer: 'return=representation' }),
        body: JSON.stringify({ error_message: 'test écriture', context: marker, url: 'run-tests' }),
      });
      if (!ins.ok) throw new Error(`INSERT ${ins.status}`);
      const row = (await ins.json())[0];
      const del = await fetchT(`${url}/rest/v1/error_logs?id=eq.${row.id}`, {
        method: 'DELETE',
        headers: supaHeaders(key, { Prefer: 'return=minimal' }),
      });
      if (!del.ok) throw new Error(`DELETE ${del.status}`);
      add('supabase: écriture/suppression réelles', true, 'INSERT + DELETE OK (error_logs)', Date.now() - t);
    } catch (e) {
      add('supabase: écriture/suppression réelles', false, e.message, Date.now() - t);
    }
  } else {
    add('supabase: écriture/suppression réelles', false, 'sautée — clé service manquante');
  }

  // 4. Brevo : clé valide (GET /v3/account, aucun email envoyé)
  if (env.BREVO_API_KEY) {
    t = Date.now();
    try {
      const r = await fetchT('https://api.brevo.com/v3/account', { headers: { 'api-key': env.BREVO_API_KEY } }, 8000);
      add('brevo: clé API valide', r.ok, r.ok ? 'compte accessible' : `HTTP ${r.status}`, Date.now() - t);
    } catch (e) {
      add('brevo: clé API valide', false, e.message, Date.now() - t);
    }
  } else {
    add('brevo: clé API valide', false, 'clé manquante');
  }

  // 5. KV backups : dernier backup présent et frais
  if (env.BACKUPS) {
    try {
      const list = await env.BACKUPS.list({ prefix: 'backup:' });
      const dates = list.keys.map((k) => k.name.replace('backup:', '')).sort();
      const last = dates[dates.length - 1];
      const fresh = last && (Date.now() - new Date(last + 'T02:00:00Z').getTime()) < 2 * 86400000;
      add('backups: dernier backup < 48h', !!fresh, last ? `dernier : ${last} (${dates.length} au total)` : 'aucun backup');
    } catch (e) {
      add('backups: dernier backup < 48h', false, e.message);
    }
  } else {
    add('backups: dernier backup < 48h', false, 'binding KV absent');
  }

  // 6. File d'attente emails : pas d'échec définitif en attente
  if (isService) {
    try {
      const r = await fetchT(`${url}/rest/v1/email_queue?status=eq.failed&select=id&limit=5`, { headers: supaHeaders(key) });
      const failed = r.ok ? await r.json() : [];
      add('email_queue: aucun échec définitif', failed.length === 0, failed.length ? `${failed.length}+ emails abandonnés après 5 tentatives — voir table email_queue` : 'file saine');
    } catch (e) {
      add('email_queue: aucun échec définitif', false, e.message);
    }
  }

  const passed = tests.filter((x) => x.ok).length;
  return secureJson({
    ok: passed === tests.length,
    passed,
    total: tests.length,
    duration_ms: Date.now() - t0,
    ran_by: bySecret ? 'X-Test-Key' : (byUser.email || byUser.id),
    tests,
  });
}
