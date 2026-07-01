// ─────────────────────────────────────────────────────────────────────────────
// Résilience Supabase — point de passage unique pour tous les appels réseau.
//
// 1. createResilientFetch : injecté dans le client Supabase (global.fetch).
//    → timeout 10 s + retry exponentiel (500ms, 1s, 2s) + log en échec final,
//    pour la totalité des appels Supabase de l'app sans toucher aux call sites.
// 2. safeQuery : wrapper explicite pour les chargements critiques,
//    retourne un fallback au lieu de laisser l'UI planter.
// 3. resilientChannel : reconnexion automatique des channels Realtime.
// 4. logError : trace les erreurs critiques dans la table error_logs.
// ─────────────────────────────────────────────────────────────────────────────

const SUPA_URL = process.env.REACT_APP_SUPABASE_URL;
const SUPA_KEY = process.env.REACT_APP_SUPABASE_ANON_KEY;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ── logError : fire-and-forget vers error_logs, throttlé pour éviter le flood ──
const lastLogAt = new Map();
export function logError(message, context = '') {
  try {
    const key = `${context}|${String(message).slice(0, 80)}`;
    const now = Date.now();
    if (lastLogAt.get(key) && now - lastLogAt.get(key) < 30000) return;
    lastLogAt.set(key, now);
    if (!SUPA_URL || !SUPA_KEY) return;
    // fetch natif volontairement (pas de retry → pas de boucle si Supabase est down)
    window.fetch(`${SUPA_URL}/rest/v1/error_logs`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        apikey: SUPA_KEY,
        Authorization: `Bearer ${SUPA_KEY}`,
        Prefer: 'return=minimal',
      },
      body: JSON.stringify({
        error_message: String(message).slice(0, 2000),
        context: String(context).slice(0, 500),
        user_agent: navigator.userAgent,
        url: window.location.href,
      }),
    }).catch(() => {});
  } catch {
    /* le logging ne doit jamais faire planter l'app */
  }
}

// ── Fetch résilient (timeout + retry) ────────────────────────────────────────
// Écritures : retry uniquement sur erreur réseau franche ou 429/503 (risque de
// doublon quasi nul). Lectures : retry aussi sur 5xx et timeout.
const RETRY_READ = new Set([408, 425, 429, 500, 502, 503, 504]);
const RETRY_WRITE = new Set([429, 503]);

export function createResilientFetch({ retries = 3, timeoutMs = 10000, onFinalError } = {}) {
  return async function resilientFetch(url, options = {}) {
    const method = (options.method || 'GET').toUpperCase();
    const isRead = method === 'GET' || method === 'HEAD';
    const retryOn = isRead ? RETRY_READ : RETRY_WRITE;
    let delay = 500;
    let lastErr;
    for (let attempt = 0; attempt <= retries; attempt++) {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), timeoutMs);
      try {
        const res = await fetch(url, { ...options, signal: options.signal || ctrl.signal });
        clearTimeout(timer);
        if (attempt === retries || !retryOn.has(res.status)) return res;
        lastErr = new Error(`HTTP ${res.status} sur ${method}`);
      } catch (e) {
        clearTimeout(timer);
        lastErr = e;
        if (options.signal && options.signal.aborted) throw e; // abort volontaire du caller
        // timeout sur écriture : la requête a pu aboutir côté serveur → pas de retry
        if (!isRead && e.name === 'AbortError') break;
        if (attempt === retries) break;
      }
      await sleep(delay);
      delay *= 2;
    }
    if (onFinalError) onFinalError(lastErr, url);
    throw lastErr;
  };
}

// ── safeQuery : requête critique avec fallback ────────────────────────────────
// Usage : const { data, error } = await safeQuery(() => supabase.from('x').select(), { fallback: [] })
const isTransient = (err) => {
  const m = String((err && err.message) || err || '').toLowerCase();
  return m.includes('failed to fetch') || m.includes('network') || m.includes('timeout') || m.includes('abort') || m.includes('load failed');
};

export async function safeQuery(queryFn, { retries = 3, fallback = null, context = '' } = {}) {
  let delay = 500;
  for (let attempt = 0; ; attempt++) {
    try {
      const res = await queryFn();
      if (!res || !res.error) return res || { data: fallback, error: null };
      if (attempt >= retries || !isTransient(res.error)) {
        logError(res.error.message || JSON.stringify(res.error), context || 'safeQuery');
        return { ...res, data: res.data != null ? res.data : fallback };
      }
    } catch (e) {
      if (attempt >= retries) {
        logError((e && e.message) || String(e), context || 'safeQuery');
        return { data: fallback, error: e };
      }
    }
    await sleep(delay);
    delay = Math.min(delay * 2, 4000);
  }
}

// ── resilientChannel : Realtime avec reconnexion auto (backoff expo, max 30 s) ─
// Usage dans un useEffect : return resilientChannel(supabase, 'nom', ch => ch.on(...).on(...))
export function resilientChannel(client, name, configure) {
  let ch = null;
  let attempts = 0;
  let disposed = false;
  let timer = null;

  const connect = () => {
    if (disposed) return;
    ch = configure(client.channel(name));
    ch.subscribe((status) => {
      if (disposed) return;
      if (status === 'SUBSCRIBED') {
        attempts = 0;
        return;
      }
      if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT' || status === 'CLOSED') {
        try { client.removeChannel(ch); } catch { /* déjà retiré */ }
        ch = null;
        const delay = Math.min(30000, 1000 * 2 ** attempts);
        attempts += 1;
        if (attempts > 1) logError(`Realtime ${name} déconnecté (${status}), reconnexion dans ${delay}ms`, 'realtime');
        clearTimeout(timer);
        timer = setTimeout(connect, delay);
      }
    });
  };

  connect();

  return () => {
    disposed = true;
    clearTimeout(timer);
    if (ch) {
      try { client.removeChannel(ch); } catch { /* noop */ }
    }
  };
}
