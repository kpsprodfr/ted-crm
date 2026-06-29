export async function onRequestPost(context) {
  const { to, message } = await context.request.json();
  const apiKey = context.env.BREVO_API_KEY;
  if (!apiKey) {
    return Response.json({ success: false, error: 'BREVO_API_KEY manquante dans CF Pages env vars' }, { status: 500 });
  }
  const numeroNettoye = to.replace(/[\s.\-()]/g, '').replace(/^0/, '+33');

  const res = await fetch('https://api.brevo.com/v3/transactionalSMS/sms', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'api-key': apiKey },
    body: JSON.stringify({
      sender: 'LE TED',
      recipient: numeroNettoye,
      content: message,
      type: 'transactional',
      unicodeEnabled: true
    })
  });
  const data = await res.json();
  return Response.json({ success: res.ok, status: res.status, data });
}
