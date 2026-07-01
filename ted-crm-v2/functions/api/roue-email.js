export async function onRequestPost(context) {
  const { env, request } = context;
  const body = await request.json();
  const { type, to_email, to_prenom, to_nom, recompense, emoji, date_venue, heure_venue, gain_id } = body;

  if (!to_email) return new Response(JSON.stringify({ error: 'to_email requis' }), { status: 400 });

  const BREVO_KEY = env.BREVO_API_KEY;
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

  const htmlContent = `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:32px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">

        <!-- HEADER -->
        <tr><td style="background:#111;padding:28px 40px;text-align:center;">
          <div style="font-size:13px;font-weight:700;letter-spacing:2px;color:#E8C547;text-transform:uppercase;margin-bottom:4px;">Grand Jeux du</div>
          <div style="font-size:32px;font-weight:900;color:#fff;letter-spacing:-1px;">TED</div>
        </td></tr>

        <!-- CORPS -->
        <tr><td style="padding:36px 40px;color:#333;font-size:15px;line-height:1.8;">
          ${htmlCorps}
        </td></tr>

        <!-- FOOTER -->
        <tr><td style="background:#111;padding:24px 40px;text-align:center;">
          <p style="margin:0 0 6px;font-size:13px;font-weight:700;color:#E8C547;">Le TED — Restaurant &amp; Club</p>
          <p style="margin:0 0 4px;font-size:12px;color:#aaa;">5 Rue Professeur Rochaix, 69008 Lyon</p>
          <p style="margin:0 0 12px;font-size:12px;color:#aaa;">📞 04 72 02 20 20</p>
          <p style="margin:0;font-size:11px;color:#666;">Vous recevez cet email car vous avez participé au Grand Jeux du TED.<br>
          <a href="mailto:contact@le-ted.fr?subject=Désinscription" style="color:#888;text-decoration:underline;">Se désinscrire</a></p>
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
      sender: { name: 'Le TED', email: 'contact@le-ted.fr' },
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
