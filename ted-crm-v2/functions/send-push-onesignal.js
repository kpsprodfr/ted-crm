// Push OneSignal (nouvelle réservation) — appelé par le CRM.
// Protégé : JWT Supabase requis + origine + rate limit.
import { fetchT, verifyUser, guard, secureJson } from './_utils.js';

export async function onRequestPost(context) {
  const { env, request } = context;

  const blocked = await guard(env, request, { limit: 20, bucket: 'push' });
  if (blocked) return blocked;

  const user = await verifyUser(env, request);
  if (!user) return secureJson({ success: false, error: 'Non autorisé' }, { status: 401 });

  let payload;
  try { payload = await request.json(); } catch { return secureJson({ success: false, error: 'JSON invalide' }, { status: 400 }); }

  let title = payload.title;
  let body = payload.body;

  if (!title && payload.record) {
    title = '📅 Nouvelle réservation !';
    body = `${payload.record.nb_personnes} pers. · ${payload.record.heure || ''} · ${payload.record.service === 'midi' ? 'Déjeuner' : 'Dîner'}`;
  }
  if (!title || !body) return secureJson({ success: false, error: 'title et body requis' }, { status: 400 });

  try {
    const res = await fetchT('https://api.onesignal.com/notifications', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Key ${env.ONESIGNAL_REST_API_KEY}` },
      body: JSON.stringify({ app_id: '87b29550-ffb0-412a-9682-05fdace514fc', included_segments: ['All'], headings: { en: String(title).slice(0, 120) }, contents: { en: String(body).slice(0, 300) } }),
    });
    const data = await res.json().catch(() => ({}));
    return secureJson({ success: res.ok, data });
  } catch (e) {
    return secureJson({ success: false, error: e.message }, { status: 502 });
  }
}
