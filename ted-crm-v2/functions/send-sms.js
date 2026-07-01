// Envoi de SMS depuis le CRM (campagnes et transactionnel).
// Protégé : JWT Supabase requis + origine + rate limit.
// La clé Brevo ne quitte jamais le serveur (plus d'appel direct navigateur → Brevo).
import { fetchT, verifyUser, guard, secureJson, isValidTelFR } from './_utils.js';

export async function onRequestPost(context) {
  const { env, request } = context;

  const blocked = await guard(env, request, { limit: 60, bucket: 'send-sms' });
  if (blocked) return blocked;

  const user = await verifyUser(env, request);
  if (!user) return secureJson({ success: false, error: 'Non autorisé' }, { status: 401 });

  const apiKey = env.BREVO_API_KEY;
  if (!apiKey) {
    return secureJson({ success: false, error: 'BREVO_API_KEY manquante dans CF Pages env vars' }, { status: 500 });
  }

  let body;
  try { body = await request.json(); } catch { return secureJson({ success: false, error: 'JSON invalide' }, { status: 400 }); }
  const { to, message, type } = body || {};
  if (!to || !message) return secureJson({ success: false, error: 'to et message requis' }, { status: 400 });
  if (!isValidTelFR(to)) return secureJson({ success: false, error: 'Numéro invalide (format FR attendu)' }, { status: 400 });
  if (String(message).length > 640) return secureJson({ success: false, error: 'Message trop long (max 640 caractères)' }, { status: 400 });

  const numeroNettoye = String(to).replace(/[\s.\-()]/g, '').replace(/^0/, '+33');
  const smsType = type === 'marketing' ? 'marketing' : 'transactional';

  try {
    const res = await fetchT('https://api.brevo.com/v3/transactionalSMS/sms', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'api-key': apiKey },
      body: JSON.stringify({
        sender: 'LE TED',
        recipient: numeroNettoye,
        content: String(message),
        type: smsType,
        unicodeEnabled: true,
      }),
    });
    const data = await res.json().catch(() => ({}));
    return secureJson({ success: res.ok, status: res.status, data });
  } catch (e) {
    return secureJson({ success: false, error: e.message }, { status: 502 });
  }
}
