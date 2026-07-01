import { sendBrevoEmail, queueEmail } from './_utils.js';

export async function onRequestPost(context) {
  const { env, request } = context;
  let body;
  try {
    body = await request.json();
  } catch {
    return Response.json({ success: false, error: 'JSON invalide' }, { status: 400 });
  }
  const { to, toName, subject, html } = body || {};
  if (!to || !subject || !html) {
    return Response.json({ success: false, error: 'to, subject et html requis' }, { status: 400 });
  }

  const result = await sendBrevoEmail(env, { to_email: to, to_name: toName, subject, html });
  if (result.ok) {
    return Response.json({ success: true });
  }

  // Échec Brevo (clé manquante, quota, panne, timeout) → l'email part en file
  // d'attente et sera repris par /api/process-email-queue. Rien n'est perdu.
  console.error('[send-email] échec Brevo:', result.status, result.detail);
  const queued = await queueEmail(env, {
    to_email: to,
    to_name: toName,
    subject,
    html,
    error_message: `Brevo ${result.status}: ${result.detail}`,
  });
  if (queued) {
    return Response.json({ success: true, queued: true, info: 'Envoi différé (file d’attente)' });
  }
  return Response.json({ success: false, error: `Brevo ${result.status}: ${result.detail}` }, { status: 502 });
}
