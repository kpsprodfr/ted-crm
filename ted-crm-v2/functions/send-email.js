export async function onRequestPost(context) {
  console.log('BREVO_API_KEY value:', context.env.BREVO_API_KEY ? 'PRESENTE - ' + context.env.BREVO_API_KEY.substring(0, 20) : 'ABSENTE')
  const { to, toName, subject, html } = await context.request.json();
  const apiKey = context.env.BREVO_API_KEY;

  if (!apiKey) {
    console.error('Brevo response status: BREVO_API_KEY manquante dans les variables Cloudflare');
    return Response.json({ success: false, error: 'BREVO_API_KEY non définie' }, { status: 500 });
  }

  const res = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify({
      sender: { name: 'Le TED', email: 'com.astegal@gmail.com' },
      to: [{ email: to, name: toName }],
      subject,
      htmlContent: html
    })
  });

  const data = await res.json();
  console.log('Brevo response status:', res.status);
  console.log('Brevo response data:', JSON.stringify(data));
  return Response.json({ success: res.ok, data });
}
