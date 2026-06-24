export async function onRequestPost(context) {
  const { title, body } = await context.request.json();
  const supabaseUrl = context.env.REACT_APP_SUPABASE_URL;
  const supabaseKey = context.env.REACT_APP_SUPABASE_ANON_KEY;
  const fcmKey = context.env.FCM_SERVER_KEY;

  // Récupère tous les tokens FCM
  const res = await fetch(`${supabaseUrl}/rest/v1/fcm_tokens?select=token`, {
    headers: {
      'apikey': supabaseKey,
      'Authorization': `Bearer ${supabaseKey}`
    }
  });
  const tokens = await res.json();

  if (!Array.isArray(tokens) || tokens.length === 0) {
    return Response.json({ success: true, sent: 0 });
  }

  // Envoie le push à chaque token via FCM HTTP v1
  let sent = 0;
  for (const { token } of tokens) {
    const r = await fetch('https://fcm.googleapis.com/fcm/send', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `key=${fcmKey}`
      },
      body: JSON.stringify({
        to: token,
        notification: { title, body },
        webpush: {
          notification: {
            icon: 'https://ted-crm.pages.dev/favicon.png',
            badge: 'https://ted-crm.pages.dev/favicon.png',
            vibrate: [200, 100, 200]
          }
        }
      })
    });
    if (r.ok) sent++;
  }

  return Response.json({ success: true, sent });
}
