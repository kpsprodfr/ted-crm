export async function onRequestPost(context) {
  const { env, request } = context;
  const body = await request.json();
  const { type, to_email, to_prenom, to_nom, recompense, emoji, date_venue, heure_venue, gain_id } = body;

  if (!to_email) return new Response(JSON.stringify({ error: 'to_email requis' }), { status: 400 });

  const BREVO_KEY = env.BREVO_API_KEY;
  if (!BREVO_KEY) {
    console.error('[roue-email] BREVO_API_KEY manquante dans les variables Cloudflare');
    return new Response(JSON.stringify({ error: 'BREVO_API_KEY non définie' }), { status: 500 });
  }
  const SUPA_URL = env.SUPABASE_URL || 'https://mwpfaytccypvdrgapptk.supabase.co';
  const SUPA_KEY = env.SUPABASE_KEY || env.SUPABASE_ANON_KEY || 'sb_publishable_4-uVtQtXd0jLGkNAFsx4yw_ni17DzN_';

  const objetCle = type === 'email1' ? 'roue_email1_objet' : 'roue_email2_objet';
  const corpsCle = type === 'email1' ? 'roue_email1_corps' : 'roue_email2_corps';

  const paramsRes = await fetch(
    `${SUPA_URL}/rest/v1/roue_config?cle=in.(${objetCle},${corpsCle})&select=cle,valeur`,
    { headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` } }
  );
  const paramsData = paramsRes.ok ? await paramsRes.json() : [];
  const p = {};
  (paramsData || []).forEach(r => { p[r.cle] = r.valeur; });

  const defaultObjet1 = '🎉 {prenom}, vous avez gagné {emoji} {recompense} au Grand Jeux du TED !';
  const defaultObjet2 = '🥂 Votre {recompense} vous attend au TED !';
  const defaultCorps1 = `Bonjour {prenom},

Félicitations ! Vous remportez : {emoji} {recompense}

Votre récompense vous attend au TED le {date}.

Pour en profiter, il vous suffit de venir accompagné(e) d'au moins 4 personnes, soit une table de 5 personnes minimum, autour d'un repas au TED.

À votre arrivée, présentez simplement cet e-mail à notre équipe : votre récompense vous sera offerte. 🥂

Une belle occasion de réunir vos proches et de célébrer cette victoire comme il se doit !

📞 Réservation par téléphone : 04 72 02 20 20
🔗 Réservation en ligne : https://ted-crm.pages.dev/reserver.html

Nous avons hâte de vous accueillir.
À très bientôt,
L'équipe du TED 🦁`;

  let objet = p[objetCle] || (type === 'email1' ? defaultObjet1 : defaultObjet2);
  let corps = p[corpsCle] || (type === 'email1' ? defaultCorps1 : '');

  const dateVenue = date_venue ? new Date(date_venue + 'T00:00:00') : null;
  const joursFR = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];

  const replace = str => str
    .replace(/{prenom}/g, to_prenom || '')
    .replace(/{nom}/g, to_nom || '')
    .replace(/{recompense}/g, recompense || '')
    .replace(/{emoji}/g, emoji || '')
    .replace(/{date}/g, dateVenue ? dateVenue.toLocaleDateString('fr-FR') : '')
    .replace(/{jour}/g, dateVenue ? joursFR[dateVenue.getDay()] : '')
    .replace(/{heure}/g, heure_venue ? String(heure_venue).slice(0, 5) : '');

  objet = replace(objet);
  const htmlCorps = replace(corps).replace(/\n/g, '<br>');

  const recompenseEmoji = replace('{emoji}');
  const recompenseNom = replace('{recompense}');

  const preheader = replace('Félicitations {prenom}, votre récompense vous attend au TED !');

  const htmlContent = `<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta name="x-apple-disable-message-reformatting">
</head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">
  <!-- Preheader (invisible, aperçu Gmail) -->
  <div style="display:none;max-height:0;overflow:hidden;font-size:1px;color:#f5f5f5;">${preheader}&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;</div>
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:32px 16px;">
    <tr><td align="center">
      <table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#111111;border-radius:12px 12px 0 0;overflow:hidden;">
        <!-- HEADER -->
        <tr><td style="background:#111111;padding:28px 24px;text-align:center;border-bottom:4px solid #E8C547;">
          <img src="https://www.leted.fr/wp-content/uploads/2024/01/Logo-Le-TED.png" alt="Le TED" style="height:60px;margin-bottom:12px;display:block;margin-left:auto;margin-right:auto;" />
          <div style="font-size:22px;font-weight:900;color:#E8C547;letter-spacing:2px;text-transform:uppercase;">Grand Jeux du TED</div>
          <div style="font-size:12px;color:#888;letter-spacing:2px;margin-top:4px;text-transform:uppercase;">Restaurant &amp; Club · Lyon</div>
        </td></tr>
      </table>
      <table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#ffffff;">
        <!-- CORPS -->
        <tr><td style="padding:32px 24px;color:#111111;font-size:15px;line-height:1.8;font-family:Arial,sans-serif;">
          ${htmlCorps}
          <!-- Encadré récompense -->
          <div style="background:#FFF8DC;border-left:4px solid #E8C547;border-radius:0 8px 8px 0;padding:20px 24px;margin:24px 0;">
            <div style="font-size:12px;font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Votre récompense</div>
            <div style="font-size:28px;margin-bottom:6px;">${recompenseEmoji}</div>
            <div style="font-size:18px;font-weight:800;color:#111111;">${recompenseNom}</div>
          </div>
          <div style="text-align:center;margin-top:28px;">
            <a href="https://ted-crm.pages.dev/reserver.html" style="display:inline-block;background:#E8C547;color:#111111;text-decoration:none;padding:14px 32px;border-radius:10px;font-weight:800;font-size:15px;font-family:Arial,sans-serif;">📅 Réserver ma table</a>
          </div>
        </td></tr>
      </table>
      <table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#111111;border-radius:0 0 12px 12px;overflow:hidden;">
        <!-- FOOTER -->
        <tr><td style="padding:24px;text-align:center;">
          <p style="margin:0 0 6px;font-size:14px;font-weight:700;color:#E8C547;">Le TED — Restaurant &amp; Club</p>
          <p style="margin:0 0 4px;font-size:12px;color:#aaa;">5 Rue Professeur Rochaix, 69003 Lyon</p>
          <p style="margin:0 0 12px;font-size:12px;color:#aaa;">📞 04 72 02 20 20</p>
          <p style="margin:0 0 8px;font-size:12px;">
            <a href="https://ted-crm.pages.dev/reserver.html" style="color:#E8C547;text-decoration:none;font-weight:700;">Réserver en ligne</a>
          </p>
          <p style="margin:0;font-size:11px;color:#555;">Vous recevez cet email car vous avez participé au Grand Jeux du TED.<br>
          <a href="mailto:contact@le-ted.fr?subject=Désinscription" style="color:#666;text-decoration:underline;">Se désinscrire</a></p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

  const brevoRes = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: { 'api-key': BREVO_KEY, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      sender: { name: 'Le TED', email: 'com.astegal@gmail.com' },
      to: [{ email: to_email, name: `${to_prenom || ''} ${to_nom || ''}`.trim() }],
      subject: objet,
      htmlContent,
    }),
  });

  if (!brevoRes.ok) {
    const err = await brevoRes.text();
    return new Response(JSON.stringify({ error: 'Brevo error', detail: err }), { status: 500 });
  }

  if (gain_id) {
    const updateField = type === 'email1' ? { email1_envoye: true } : { email2_envoye: true };
    await fetch(`${SUPA_URL}/rest/v1/roue_gains?id=eq.${gain_id}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json', apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}`, Prefer: 'return=minimal' },
      body: JSON.stringify(updateField),
    });
  }

  return new Response(JSON.stringify({ ok: true }), { status: 200 });
}
