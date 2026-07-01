// Utilitaires partagés des Cloudflare Pages Functions.
// (préfixe "_" = non routé par Pages, importable par les functions)

// Tables couvertes par le backup quotidien et la restauration
export const BACKUP_TABLES = [
  'clients', 'reservations', 'roue_gains', 'roue_recompenses', 'roue_config',
  'parametres', 'menu_produits', 'menu_categories', 'menu_cartes',
  'menu_soirees', 'menu_plat_jour', 'menu_origines',
];

// ── fetch avec timeout 10 s ──────────────────────────────────────────────────
export async function fetchT(url, options = {}, timeoutMs = 10000) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: ctrl.signal });
  } finally {
    clearTimeout(timer);
  }
}

// ── Accès Supabase depuis les functions ──────────────────────────────────────
export function getSupa(env) {
  const url = env.SUPABASE_URL || 'https://mwpfaytccypvdrgapptk.supabase.co';
  const serviceKey = env.SUPABASE_SERVICE_ROLE_KEY || env.SUPABASE_SERVICE_KEY || null;
  const anonKey = env.SUPABASE_ANON_KEY || env.SUPABASE_KEY || 'sb_publishable_4-uVtQtXd0jLGkNAFsx4yw_ni17DzN_';
  const key = serviceKey || anonKey;
  return { url, key, isService: !!serviceKey };
}

export function supaHeaders(key, extra = {}) {
  return { 'Content-Type': 'application/json', apikey: key, Authorization: `Bearer ${key}`, ...extra };
}

// ── File d'attente email : aucun email ne doit être perdu ────────────────────
// Insère l'email rendu dans email_queue (policy RLS : INSERT autorisé en anon).
// Le cron /api/process-email-queue reprendra les 'pending'.
export async function queueEmail(env, { to_email, to_name, subject, html, error_message }) {
  try {
    const { url, key } = getSupa(env);
    const res = await fetchT(`${url}/rest/v1/email_queue`, {
      method: 'POST',
      headers: supaHeaders(key, { Prefer: 'return=minimal' }),
      body: JSON.stringify({
        to_email,
        to_name: to_name || null,
        subject: subject || null,
        html: html || null,
        status: 'pending',
        attempts: 0,
        error_message: (error_message || '').slice(0, 1000) || null,
      }),
    });
    return res.ok;
  } catch (e) {
    console.error('[queueEmail] échec mise en file:', e.message);
    return false;
  }
}

// ── Headers de sécurité sur toutes les réponses ──────────────────────────────
export const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'Strict-Transport-Security': 'max-age=31536000; includeSubDomains',
  'Content-Security-Policy': "default-src 'none'",
  'Referrer-Policy': 'no-referrer',
};

export function secureJson(data, init = {}) {
  return Response.json(data, { ...init, headers: { ...SECURITY_HEADERS, ...(init.headers || {}) } });
}

// ── Vérification d'origine ───────────────────────────────────────────────────
// Les POST des navigateurs portent toujours un header Origin : on n'accepte que
// le site (+ previews Pages et localhost pour le dev). Les GET server-to-server
// (pg_cron) n'ont pas d'Origin → autorisés.
export function checkOrigin(request) {
  const src = request.headers.get('Origin') || request.headers.get('Referer') || '';
  if (!src) return request.method === 'GET' || request.method === 'HEAD';
  try {
    const h = new URL(src).hostname;
    return h === 'ted-crm.pages.dev' || h.endsWith('.ted-crm.pages.dev') || h === 'localhost' || h === '127.0.0.1';
  } catch {
    return false;
  }
}

// ── Rate limiting simple : 10 req/min/IP via KV ──────────────────────────────
// Approximatif (KV est à cohérence éventuelle) mais suffisant pour stopper les abus.
// Sans binding KV : pas de rate limit (dégradé, signalé par /api/health).
export async function rateLimit(env, request, { limit = 10, windowSec = 60, bucket = 'fn' } = {}) {
  if (!env.BACKUPS) return true;
  try {
    const ip = request.headers.get('CF-Connecting-IP') || 'unknown';
    const windowId = Math.floor(Date.now() / (windowSec * 1000));
    const kvKey = `rl:${bucket}:${ip}:${windowId}`;
    const current = parseInt((await env.BACKUPS.get(kvKey)) || '0', 10);
    if (current >= limit) return false;
    await env.BACKUPS.put(kvKey, String(current + 1), { expirationTtl: windowSec * 2 });
    return true;
  } catch {
    return true;
  }
}

// Garde standard pour les endpoints exposés : origine + rate limit.
// Retourne une Response d'erreur, ou null si tout est OK.
export async function guard(env, request, { limit = 10, bucket = 'fn' } = {}) {
  if (!checkOrigin(request)) return secureJson({ error: 'Origine non autorisée' }, { status: 403 });
  if (!(await rateLimit(env, request, { limit, bucket }))) {
    return secureJson({ error: 'Trop de requêtes, réessayez dans une minute' }, { status: 429 });
  }
  return null;
}

// ── Échappement HTML (anti-injection dans les emails) ────────────────────────
export function escapeHtml(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ── Validations serveur (jamais confiance au client) ─────────────────────────
export const isValidEmail = (v) => typeof v === 'string' && v.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v);
export const isValidTelFR = (v) => typeof v === 'string' && /^(\+33|0)[1-9]\d{8}$/.test(String(v).replace(/[\s.\-()]/g, ''));

// ── Vérification d'un utilisateur CRM connecté (JWT Supabase Auth) ──────────
// Les actions sensibles (backup manuel, restauration) exigent un JWT valide
// d'un utilisateur du CRM : Authorization: Bearer <access_token>.
export async function verifyUser(env, request) {
  const auth = request.headers.get('Authorization') || '';
  if (!auth.startsWith('Bearer ')) return null;
  const url = env.SUPABASE_URL || 'https://mwpfaytccypvdrgapptk.supabase.co';
  const anonKey = env.SUPABASE_ANON_KEY || env.SUPABASE_KEY || 'sb_publishable_4-uVtQtXd0jLGkNAFsx4yw_ni17DzN_';
  try {
    const res = await fetchT(`${url}/auth/v1/user`, { headers: { apikey: anonKey, Authorization: auth } });
    if (!res.ok) return null;
    const user = await res.json();
    return user && user.id ? user : null;
  } catch {
    return null;
  }
}

export async function sha256Hex(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

// ── Envoi Brevo mutualisé ────────────────────────────────────────────────────
// Retourne { ok, status, detail }. Ne jette jamais.
export async function sendBrevoEmail(env, { to_email, to_name, subject, html }) {
  const apiKey = env.BREVO_API_KEY;
  if (!apiKey) return { ok: false, status: 0, detail: 'BREVO_API_KEY manquante' };
  try {
    const res = await fetchT('https://api.brevo.com/v3/smtp/email', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'api-key': apiKey },
      body: JSON.stringify({
        sender: { name: 'Le TED', email: 'com.astegal@gmail.com' },
        to: [{ email: to_email, name: to_name || 'Client' }],
        subject,
        htmlContent: html,
      }),
    });
    const detail = res.ok ? '' : (await res.text()).slice(0, 500);
    return { ok: res.ok, status: res.status, detail };
  } catch (e) {
    return { ok: false, status: 0, detail: e.message };
  }
}
