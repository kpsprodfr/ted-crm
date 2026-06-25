export async function onRequestPost(context) {
  const { title, body } = await context.request.json();
  const supabaseUrl = context.env.REACT_APP_SUPABASE_URL;
  const supabaseKey = context.env.REACT_APP_SUPABASE_ANON_KEY;
  const serviceAccount = JSON.parse(context.env.FIREBASE_SERVICE_ACCOUNT);

  // 1. Récupère tous les tokens FCM depuis Supabase
  const res = await fetch(`${supabaseUrl}/rest/v1/fcm_tokens?select=token`, {
    headers: {
      'apikey': supabaseKey,
      'Authorization': `Bearer ${supabaseKey}`
    }
  });
  const tokens = await res.json();
  if (!tokens?.length) return Response.json({ success: true, sent: 0 });

  // 2. Génère un JWT pour OAuth2 Google
  function str2ab(str) {
    const binary = atob(str);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
  }

  async function getAccessToken(sa) {
    const now = Math.floor(Date.now() / 1000);
    const encode = obj => btoa(JSON.stringify(obj))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');

    const header = { alg: 'RS256', typ: 'JWT' };
    const payload = {
      iss: sa.client_email,
      scope: 'https://www.googleapis.com/auth/firebase.messaging',
      aud: 'https://oauth2.googleapis.com/token',
      iat: now,
      exp: now + 3600
    };

    const signingInput = `${encode(header)}.${encode(payload)}`;

    const privateKey = await crypto.subtle.importKey(
      'pkcs8',
      str2ab(sa.private_key
        .replace('-----BEGIN PRIVATE KEY-----', '')
        .replace('-----END PRIVATE KEY-----', '')
        .replace(/\n/g, '')),
      { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
      false,
      ['sign']
    );

    const signature = await crypto.subtle.sign(
      'RSASSA-PKCS1-v1_5',
      privateKey,
      new TextEncoder().encode(signingInput)
    );

    const jwt = `${signingInput}.${btoa(String.fromCharCode(...new Uint8Array(signature)))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '')}`;

    const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
    });
    const tokenData = await tokenRes.json();
    return tokenData.access_token;
  }

  const accessToken = await getAccessToken(serviceAccount);
  const projectId = serviceAccount.project_id;

  // 3. Envoie à chaque token via FCM v1
  let sent = 0;
  const errors = [];
  for (const { token } of tokens) {
    const fcmRes = await fetch(
      `https://fcm.googleapis.com/v1/projects/${projectId}/messages:send`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${accessToken}`
        },
        body: JSON.stringify({
          message: {
            token,
            notification: { title, body },
            webpush: {
              notification: {
                icon: 'https://ted-crm.pages.dev/favicon.png',
                badge: 'https://ted-crm.pages.dev/favicon.png',
                vibrate: [200, 100, 200],
                tag: 'nouvelle-resa',
                renotify: true
              },
              fcm_options: {
                link: 'https://ted-crm.pages.dev'
              }
            }
          }
        })
      }
    );
    if (fcmRes.ok) {
      sent++;
    } else {
      const err = await fcmRes.json();
      errors.push({ token: token.slice(-8), error: err?.error?.message });
    }
  }

  return Response.json({ success: true, sent, errors });
}
