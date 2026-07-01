// Envoi d'email depuis le CRM (confirmations de réservation, communications).
// Protégé : JWT Supabase requis (utilisateur CRM connecté) + origine + rate limit.
// En cas d'échec Brevo, l'email part en file d'attente (email_queue) — rien n'est perdu.
import { sendBrevoEmail, queueEmail, verifyUser, guard, secureJson, isValidEmail } from './_utils.js';

export async function onRequestPost(context) {
  const { env, request } = context;

  const blocked = await guard(env, request, { limit: 30, bucket: 'send-email' });
  if (blocked) return blocked;

  const user = await verifyUser(env, request);
  if (!user) return secureJson({ success: false, error: 'Non autorisé' }, { status: 401 });

  let body;
  try {
    body = await request.json();
  } catch {
    return secureJson({ success: false, error: 'JSON invalide' }, { status: 400 });
  }
  const { to, toName, subject, html } = body || {};
  if (!to || !subject || !html) {
    return secureJson({ success: false, error: 'to, subject et html requis' }, { status: 400 });
  }
  if (!isValidEmail(to)) return secureJson({ success: false, error: 'Adresse email invalide' }, { status: 400 });

  const result = await sendBrevoEmail(env, { to_email: to, to_name: toName, subject, html });
  if (result.ok) {
    return secureJson({ success: true });
  }

  // Échec Brevo (quota, panne, timeout) → file d'attente, reprise automatique.
  console.error('[send-email] échec Brevo:', result.status, result.detail);
  const queued = await queueEmail(env, {
    to_email: to,
    to_name: toName,
    subject,
    html,
    error_message: `Brevo ${result.status}: ${result.detail}`,
  });
  if (queued) {
    return secureJson({ success: true, queued: true, info: 'Envoi différé (file d’attente)' });
  }
  return secureJson({ success: false, error: `Brevo ${result.status}: ${result.detail}` }, { status: 502 });
}
