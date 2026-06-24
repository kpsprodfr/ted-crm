export async function onRequestPost(context) {
  const { to, toName, subject, html } = await context.request.json();
  const apiKey = context.env.REACT_APP_BREVO_API_KEY;

  const res = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify({
      sender: { name: 'Le TED', email: 'reservations@leted.fr' },
      to: [{ email: to, name: toName }],
      subject,
      htmlContent: html
    })
  });

  const data = await res.json();
  return Response.json({ success: res.ok, data });
}
