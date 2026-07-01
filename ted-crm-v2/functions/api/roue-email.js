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

  const htmlContent = `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#f8f8f8;padding:20px">
  <div style="display:none;max-height:0;overflow:hidden;font-size:1px;color:#f5f5f5;">${preheader}&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;</div>

  <div style="background:#111111;padding:28px 24px;text-align:center;border-radius:12px 12px 0 0;border-bottom:4px solid #E8C547">
    <img src="https://ted-crm.pages.dev/favicon.png" alt="Le TED" style="height:60px;margin-bottom:12px" />
    <h1 style="color:#E8C547;margin:0;font-size:28px;letter-spacing:2px;font-weight:800">LE TED</h1>
    <p style="color:#888;margin:4px 0 0;font-size:13px;letter-spacing:1px">RESTAURANT &amp; CLUB — CHASSIEU</p>
  </div>

  <div style="background:#fff;padding:28px 24px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08)">
    <h2 style="color:#111;margin:0 0 8px;font-size:22px">Bonjour ${replace('{prenom}')} 🎉</h2>
    <p style="color:#444;font-size:16px;margin:0 0 24px">Vous avez remporté une récompense au <strong>Grand Jeux du TED</strong> !</p>

    <div style="background:#f9f9f9;border-left:4px solid #E8C547;padding:20px;border-radius:0 8px 8px 0;margin-bottom:24px">
      <p style="margin:0 0 10px;font-size:15px">🎁 <strong>Votre récompense :</strong> ${replace('{emoji} {recompense}')}</p>
      <p style="margin:0 0 10px;font-size:15px">📅 <strong>Disponibilité :</strong> ${dateAffichee}</p>
      <p style="margin:0;font-size:15px">📋 <strong>Conditions :</strong> Présenter ce mail à votre arrivée</p>
      ${messagePerso ? `<p style="margin:10px 0 0;font-size:14px;color:#B8960C;font-weight:600;line-height:1.5;">ℹ️ ${messagePerso.replace(/\n/g,'<br>')}</p>` : ''}
    </div>

    <div style="background:#f9f9f9;border:1.5px solid #ddd;border-radius:8px;padding:16px;margin-bottom:24px">
      <p style="margin:0 0 10px;font-size:14px;font-weight:800;color:#111">👥 Condition de retrait</p>
      <p style="margin:0;font-size:13px;color:#555;line-height:1.6">Votre récompense est valable pour une table de <strong>5 personnes minimum</strong>, autour d'un repas au TED. Elle vous sera remise à votre arrivée sur présentation de cet email.</p>
    </div>

    <div style="text-align:center;margin-bottom:24px">
      <a href="https://ted-crm.pages.dev/reserver.html" target="_blank" style="display:inline-block;background:#E8C547;color:#111;text-decoration:none;padding:14px 32px;border-radius:10px;font-weight:800;font-size:15px">📅 Réserver ma table</a>
    </div>

    <div style="background:#fff8e1;border:1.5px solid #E8C547;border-radius:8px;padding:16px;margin-bottom:24px">
      <p style="margin:0;font-size:14px;color:#555;line-height:1.6">📞 <strong>Réservation par téléphone :</strong> 04 72 02 20 20<br>Ou en ligne via le bouton ci-dessus.</p>
    </div>

    <div style="border-top:1px solid #eee;padding-top:20px;text-align:center">
      <p style="color:#111;font-weight:700;font-size:15px;margin:0 0 6px">Le TED — Restaurant &amp; Club</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📍 28 Av. des Frères Montgolfier, 69680 Chassieu</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📞 04 78 90 67 80</p>
      <p style="margin:8px 0 0;text-align:center"><a href="https://leted.fr" style="display:inline-flex;align-items:center;justify-content:center;gap:8px;text-decoration:none;color:#111;font-size:15px;font-weight:700"><img src="https://ted-crm.pages.dev/favicon.png" alt="TED" style="height:24px;width:24px;vertical-align:middle" />leted.fr</a></p>
    </div>
    <p style="text-align:center;color:#bbb;font-size:12px;margin-top:20px">Nous avons hâte de vous accueillir ! 🎉</p>
  </div>
</div>`;

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
