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
    `${SUPA_URL}/rest/v1/roue_config?cle=in.(${objetCle},${corpsCle},roue_email_date,roue_email_date_fin,roue_email_date_mode,roue_email_message)&select=cle,valeur`,
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
  const messagePerso = p['roue_email_message'] || '';

  // Date depuis les paramètres CRM ou depuis le body de la requête
  const dateMode = p['roue_email_date_mode'] || 'precise';
  const dateDebutIso = p['roue_email_date'] || date_venue || null;
  const dateFinIso = p['roue_email_date_fin'] || null;
  const fmtDateFR = iso => iso ? new Date(iso + 'T00:00:00').toLocaleDateString('fr-FR', { day:'numeric', month:'long', year:'numeric' }) : null;
  const dateDebutFmt = fmtDateFR(dateDebutIso);
  const dateFinFmt = fmtDateFR(dateFinIso);
  const dateAffichee = dateMode === 'periode' && dateDebutFmt && dateFinFmt
    ? `Du ${dateDebutFmt} au ${dateFinFmt}`
    : (dateDebutFmt || 'À définir par le restaurant');
  const dispoLabel = dateMode === 'periode' && dateDebutFmt && dateFinFmt
    ? `Disponible du ${dateDebutFmt} au ${dateFinFmt}`
    : (dateDebutFmt ? `Disponible à partir du ${dateDebutFmt}` : 'Disponible à partir du — À définir');

  const dateVenue = date_venue ? new Date(date_venue + 'T00:00:00') : null;
  const joursFR = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];

  const dateFormatee = dateAffichee;

  const replace = str => str
    .replace(/{prenom}/g, to_prenom || '')
    .replace(/{nom}/g, to_nom || '')
    .replace(/{recompense}/g, recompense || '')
    .replace(/{emoji}/g, emoji || '')
    .replace(/{date}/g, dateFormatee)
    .replace(/{jour}/g, dateVenue ? joursFR[dateVenue.getDay()] : '')
    .replace(/{heure}/g, heure_venue ? String(heure_venue).slice(0, 5) : '');

  objet = replace(objet);
  const htmlCorps = replace(corps).replace(/\n/g, '<br>');

  const recompenseEmoji = replace('{emoji}');
  const recompenseNom = replace('{recompense}');

  const preheader = replace('Félicitations {prenom}, votre récompense vous attend au TED !');

  const htmlContent = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="x-apple-disable-message-reformatting">
<!-- Preheader (invisible, aperçu Gmail) -->
<div style="display:none;max-height:0;overflow:hidden;font-size:1px;color:#f5f5f5;">${preheader}&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;</div>
<link href="https://fonts.googleapis.com/css2?family=Caveat:wght@700&display=swap" rel="stylesheet">
</head>
<body style="margin:0;padding:0;background:#ffffff;font-family:Arial,sans-serif;">
<div style="max-width:560px;margin:0 auto;background:#ffffff;">

  <div style="background:linear-gradient(180deg,#fff8c0 0%,#FFE033 50%,#FFC200 100%);padding:32px 24px;text-align:center;">
    <img src="https://ted-crm.pages.dev/logo-Le-TED.png" width="80" height="80" style="margin-bottom:14px;object-fit:contain;filter:brightness(0);" />
    <div style="font-family:'Caveat',cursive;color:#111;font-size:34px;font-weight:700;line-height:1.1;">Grand Jeux du <span style="font-family:Arial,sans-serif;font-weight:900;font-size:30px;letter-spacing:1px;">TED</span></div>
    <div style="color:#5a4500;font-size:11px;letter-spacing:2px;margin-top:8px;text-transform:uppercase;">Restaurant &amp; Club</div>
  </div>

  <div style="background:#ffffff;padding:32px 28px;">
    <p style="color:#333;font-size:15px;margin:0 0 28px;line-height:1.6;">
      Bonjour <strong style="color:#111;font-size:16px;">${replace('{prenom} {nom}')}</strong>,
    </p>

    <div style="background:rgba(232,197,71,0.12);border:1.5px solid rgba(232,197,71,0.4);border-radius:14px;padding:28px 24px;text-align:center;margin-bottom:24px;">
      <div style="color:#111;font-size:28px;font-weight:700;margin-bottom:16px;">Votre récompense</div>
      <div style="background:#fff;border:1.5px solid rgba(232,197,71,0.5);border-radius:10px;padding:16px 20px;margin-bottom:16px;display:inline-block;min-width:80%;">
        <div style="color:#B8960C;font-size:18px;font-weight:700;">${replace('{emoji} {recompense}')}</div>
      </div>
      <div style="font-size:42px;margin-bottom:12px;">🥳</div>
      <div style="color:#888;font-size:12px;font-style:italic;">${dispoLabel}</div>
      ${messagePerso ? `<div style="color:#888;font-size:12px;font-style:italic;margin-top:8px;line-height:1.5;">${messagePerso.replace(/\n/g,'<br>')}</div>` : ''}
    </div>

    <div style="border:1.5px solid #E8C547;border-radius:12px;padding:22px 24px;margin-bottom:28px;">
      <div style="color:#111;font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:18px;">Conditions de retrait</div>

      <div style="display:flex;gap:14px;margin-bottom:16px;align-items:flex-start;">
        <span style="font-size:20px;flex-shrink:0;">📋</span>
        <div>
          <div style="font-size:13px;font-weight:700;color:#111;margin-bottom:2px;">Présentation de ce mail</div>
          <div style="font-size:13px;color:#666;">À montrer à notre équipe à votre arrivée</div>
        </div>
      </div>

      <div style="display:flex;gap:14px;${messagePerso ? '' : 'margin-bottom:16px;'}align-items:flex-start;">
        <span style="font-size:20px;flex-shrink:0;">📅</span>
        <div>
          <div style="font-size:13px;font-weight:700;color:#111;margin-bottom:2px;">Date de retrait de votre cadeau</div>
          <div style="font-size:13px;color:#B8960C;font-weight:700;">${dateAffichee}</div>
          ${messagePerso ? `<div style="font-size:13px;color:#B8960C;font-weight:700;margin-top:5px;line-height:1.5;">${messagePerso.replace(/\n/g,'<br>')}</div>` : ''}
        </div>
      </div>

      <div style="display:flex;gap:14px;margin-top:16px;align-items:flex-start;">
        <span style="font-size:20px;flex-shrink:0;">👥</span>
        <div>
          <div style="font-size:13px;font-weight:700;color:#111;margin-bottom:2px;">Réservation requise</div>
          <div style="font-size:13px;color:#666;">5 personnes minimum — cette récompense accompagnera votre repas</div>
        </div>
      </div>
    </div>

    <div style="text-align:center;margin-bottom:28px;">
      <a href="https://ted-crm.pages.dev/reserver.html" style="display:inline-block;background:#E8C547;color:#111;font-weight:700;font-size:14px;padding:14px 36px;border-radius:8px;text-decoration:none;">Réserver ma table</a>
    </div>

    <p style="color:#666;font-size:13px;line-height:1.9;margin:0;border-top:0.5px solid #eee;padding-top:20px;">
      On vous attend avec impatience.<br>
      À très bientôt,<br>
      <strong style="color:#111;">L'équipe du TED</strong>
    </p>
  </div>

  <div style="background:#111;padding:22px 24px;text-align:center;">
    <div style="width:50px;height:50px;background:#fff;border-radius:50%;margin:0 auto 10px;display:flex;align-items:center;justify-content:center;"><img src="https://ted-crm.pages.dev/logo-Le-TED.png" width="36" height="36" style="object-fit:contain;" /></div>
    <div style="color:#E8C547;font-size:13px;font-weight:700;margin-bottom:4px;">Le TED — Restaurant &amp; Club</div>
    <div style="color:#888;font-size:12px;margin-bottom:4px;">28 Avenue des Frères Montgolfier, 69680 Chassieu</div>
    <div style="color:#888;font-size:12px;margin-bottom:4px;">04 78 90 67 80</div>
    <div style="margin-bottom:14px;"><a href="https://www.leted.fr" style="color:#888;font-size:12px;text-decoration:none;">www.leted.fr</a></div>
    <a href="https://ted-crm.pages.dev/reserver.html" style="color:#E8C547;font-size:12px;text-decoration:none;">Réserver en ligne</a>
    <div style="color:#555;font-size:11px;margin-top:14px;">
      Vous recevez cet email car vous avez participé au Grand Jeux du TED.<br>
      <a href="mailto:contact@le-ted.fr?subject=Désinscription" style="color:#555;">Se désinscrire</a>
    </div>
  </div>

</div>
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
