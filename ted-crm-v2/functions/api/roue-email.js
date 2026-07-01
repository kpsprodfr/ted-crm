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

  const paramsRes = await fetch(
    `${SUPA_URL}/rest/v1/roue_config?cle=in.(${objetCle},roue_email_date,roue_email_date_fin,roue_email_date_mode,roue_email_message)&select=cle,valeur`,
    { headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` } }
  );
  const paramsData = paramsRes.ok ? await paramsRes.json() : [];
  const p = {};
  (paramsData || []).forEach(r => { p[r.cle] = r.valeur; });

  const defaultObjet = '🎉 {prenom}, vous avez gagné {emoji} {recompense} au Grand Jeu du TED !';
  let objet = p[objetCle] || defaultObjet;
  const messagePerso = p['roue_email_message'] || '';

  const dateMode     = p['roue_email_date_mode'] || 'precise';
  const dateDebutIso = p['roue_email_date'] || date_venue || null;
  const dateFinIso   = p['roue_email_date_fin'] || null;
  const fmtDateFR    = iso => iso ? new Date(iso + 'T00:00:00').toLocaleDateString('fr-FR', { day:'numeric', month:'long', year:'numeric' }) : null;
  const dateDebutFmt = fmtDateFR(dateDebutIso);
  const dateFinFmt   = fmtDateFR(dateFinIso);
  const dateAffichee = dateMode === 'periode' && dateDebutFmt && dateFinFmt
    ? `${dateDebutFmt} au ${dateFinFmt}`
    : (dateDebutFmt || 'À définir');
  const dateLabel = heure_venue
    ? `${dateAffichee} · ${String(heure_venue).slice(0, 5)}`
    : dateAffichee;

  const replace = str => str
    .replace(/{prenom}/g, to_prenom || '')
    .replace(/{nom}/g,    to_nom    || '')
    .replace(/{recompense}/g, recompense || '')
    .replace(/{emoji}/g,  emoji || '')
    .replace(/{date}/g,   dateAffichee);

  objet = replace(objet);

  const preheader = `Félicitations ${to_prenom || ''}, votre récompense vous attend au TED !`;
  const serialCode = 'TED-' + Math.random().toString(36).slice(2, 7).toUpperCase();
  const logoUrl    = 'https://ted-crm.pages.dev/logo-Le-TED.png';
  const hasMessage = messagePerso.trim().length > 0;

  // Email-compatible HTML (table-based, no flex, no animations, no background-clip)
  const htmlContent = `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>Votre récompense Le TED</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Alex+Brush&display=swap');
  body, table, td { margin:0; padding:0; border:0; }
  body { background-color:#181818 !important; font-family:-apple-system,'Helvetica Neue',Arial,sans-serif; }
  img { border:0; display:block; }
  .alex { font-family:'Alex Brush',Georgia,serif; }
  /* Force dark bg sur Gmail */
  div[style*="background"] { background-color:#181818 !important; }
  u + #body { background-color:#181818 !important; }
</style>
</head>
<body id="body" style="margin:0;padding:0;background-color:#181818;" bgcolor="#181818">

<!-- Preheader -->
<div style="display:none;max-height:0;overflow:hidden;font-size:1px;color:#181818;">${preheader}&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;</div>

<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#181818" style="background-color:#181818;">
<tr><td align="center" bgcolor="#181818" style="padding:32px 16px;background-color:#181818;">

  <table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;">

    <!-- ══ HEADER ══ -->
    <tr><td bgcolor="#111111" style="background-color:#111111;border-radius:10px 10px 0 0;padding:44px 40px 32px;text-align:center;">
      <img src="${logoUrl}" width="72" height="72" alt="Le TED" style="display:block;margin:0 auto 16px;border-radius:50%;">
      <p style="margin:0;font-size:22px;font-weight:800;letter-spacing:10px;text-transform:uppercase;color:#F0A830;">LE TED</p>
      <p style="margin:6px 0 0;font-size:10px;letter-spacing:4px;color:rgba(255,255,255,0.25);text-transform:uppercase;">Restaurant &amp; Club · Chassieu</p>
    </td></tr>

    <!-- Gold bar -->
    <tr><td style="height:3px;background:linear-gradient(90deg,#c47e10,#F0A830,#ffd278,#F0A830,#c47e10);font-size:0;line-height:0;">&nbsp;</td></tr>

    <!-- ══ WINNER HERO ══ -->
    <tr><td bgcolor="#111111" style="background-color:#111111;padding:48px 40px 40px;text-align:center;">
      <p style="margin:0 0 4px;font-size:10px;font-weight:800;letter-spacing:5px;color:#F0A830;text-transform:uppercase;">&#10022; Vous avez gagné &#10022;</p>
      <h1 style="margin:14px 0 8px;font-size:36px;font-weight:700;color:#ffffff;line-height:1.15;">${to_prenom || ''} ${to_nom || ''}</h1>
      <p style="margin:0 0 40px;font-size:13px;color:rgba(255,255,255,0.45);letter-spacing:0.5px;">vous repart avec une récompense exclusive</p>

      <!-- ══ TICKET BON GAGNANT ══ -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid rgba(240,168,48,0.35);border-radius:14px;overflow:hidden;">

        <!-- Top gold bar -->
        <tr><td style="height:4px;background:linear-gradient(90deg,#c47e10,#F0A830,#ffd278,#F0A830,#c47e10);font-size:0;line-height:0;">&nbsp;</td></tr>

        <!-- Header strip -->
        <tr><td bgcolor="#1a1300" style="background-color:#1a1300;padding:12px 24px;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td style="vertical-align:middle;">
                <img src="${logoUrl}" width="20" height="20" alt="" style="display:inline-block;vertical-align:middle;opacity:0.85;margin-right:8px;">
                <span style="font-size:9px;font-weight:700;letter-spacing:3.5px;color:#F0A830;text-transform:uppercase;vertical-align:middle;">Bon Gagnant</span>
              </td>
              <td align="right" style="font-family:'Courier New',Courier,monospace;font-size:10px;color:rgba(240,168,48,0.4);letter-spacing:1px;">${serialCode}</td>
            </tr>
          </table>
        </td></tr>

        <!-- Dashed sep -->
        <tr><td style="height:1px;border-top:1px dashed rgba(240,168,48,0.2);background:#111111;font-size:0;line-height:0;">&nbsp;</td></tr>

        <!-- Body -->
        <tr><td bgcolor="#111111" style="background-color:#111111;padding:36px 28px 28px;text-align:center;">
          <p style="margin:0 0 16px;font-size:9px;letter-spacing:7px;color:rgba(240,168,48,0.4);">&#10022; &nbsp; &#10022; &nbsp; &#10022;</p>
          <p style="margin:0 0 16px;font-size:56px;line-height:1;">${emoji || '🎁'}</p>
          <p style="margin:0 0 10px;font-size:9px;font-weight:600;letter-spacing:4px;color:rgba(240,168,48,0.5);text-transform:uppercase;">Votre récompense</p>
          <p style="margin:0 0 24px;font-size:24px;font-weight:700;color:#ffffff;line-height:1.2;">${recompense || ''}</p>

          <!-- Valable pill -->
          <table cellpadding="0" cellspacing="0" border="0" align="center" style="margin-bottom:20px;">
            <tr>
              <td style="background:rgba(240,168,48,0.08);border:1px solid rgba(240,168,48,0.2);border-radius:40px;padding:8px 20px;">
                <span style="display:inline-block;width:6px;height:6px;background:#F0A830;border-radius:50%;vertical-align:middle;margin-right:8px;"></span>
                <span style="font-size:13px;color:rgba(255,255,255,0.5);vertical-align:middle;">Valable&nbsp;</span>
                <span style="font-size:13px;font-weight:700;color:#F0A830;vertical-align:middle;">${dateLabel}</span>
              </td>
            </tr>
          </table>
        </td></tr>

        <!-- Tear line -->
        <tr><td style="background:#111111;padding:0 0;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td width="12" style="background:#181818;border-radius:0 50% 50% 0;height:18px;font-size:0;"></td>
              <td style="border-top:1.5px dashed rgba(240,168,48,0.25);font-size:0;height:1px;"></td>
              <td width="12" style="background:#181818;border-radius:50% 0 0 50%;height:18px;font-size:0;"></td>
            </tr>
          </table>
        </td></tr>

        <!-- Stub footer -->
        <tr><td bgcolor="#0e0e0e" style="background-color:#0e0e0e;padding:12px 24px;text-align:center;">
          <span style="font-size:11px;color:rgba(255,255,255,0.4);text-decoration:underline;">Conditions de retrait ci-dessous &#8595;</span>
        </td></tr>

        <!-- Bottom gold bar -->
        <tr><td style="height:3px;background:linear-gradient(90deg,#c47e10,#F0A830,#ffd278,#F0A830,#c47e10);opacity:0.6;font-size:0;line-height:0;">&nbsp;</td></tr>

      </table>
      <!-- END TICKET -->

      ${hasMessage ? `
      <!-- Message perso -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:32px;">
        <tr>
          <td width="3" style="background:rgba(240,168,48,0.4);border-radius:2px;">&nbsp;</td>
          <td style="padding:16px 20px;background:rgba(255,255,255,0.03);border-radius:0 6px 6px 0;text-align:left;">
            <p style="margin:0 0 8px;font-style:italic;font-size:16px;color:rgba(255,255,255,0.6);line-height:1.7;">&laquo;&nbsp;${messagePerso.replace(/\n/g, '<br>')}&nbsp;&raquo;</p>
            <p style="margin:0;font-size:10px;letter-spacing:2.5px;color:rgba(240,168,48,0.5);text-transform:uppercase;">L'équipe du TED</p>
          </td>
        </tr>
      </table>` : ''}

    </td></tr>

    <!-- ══ CONDITIONS ══ -->
    <tr><td style="background:#FDFAF5;padding:40px 40px 36px;">
      <p style="margin:0 0 24px;font-size:20px;font-weight:700;letter-spacing:0.5px;color:#999999;text-transform:uppercase;">Conditions de retrait</p>

      <!-- Condition 1 -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-bottom:1px solid #f0ece4;margin-bottom:0;">
        <tr>
          <td width="46" valign="top" style="padding:14px 16px 14px 0;">
            <table cellpadding="0" cellspacing="0" border="0" style="background:#111111;border-radius:4px;width:38px;height:38px;">
              <tr><td align="center" valign="middle" style="font-size:18px;width:38px;height:38px;">📋</td></tr>
            </table>
          </td>
          <td valign="top" style="padding:14px 0;">
            <p style="margin:0 0 3px;font-size:14px;font-weight:700;color:#111111;">Présentation obligatoire</p>
            <p style="margin:0;font-size:12px;color:#777777;line-height:1.6;">Présentez <strong style="color:#111111;">cet email</strong> à votre arrivée au restaurant.</p>
          </td>
        </tr>
      </table>

      <!-- Condition 2 -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-bottom:1px solid #f0ece4;">
        <tr>
          <td width="46" valign="top" style="padding:14px 16px 14px 0;">
            <table cellpadding="0" cellspacing="0" border="0" style="background:#111111;border-radius:4px;width:38px;height:38px;">
              <tr><td align="center" valign="middle" style="font-size:18px;width:38px;height:38px;">📅</td></tr>
            </table>
          </td>
          <td valign="top" style="padding:14px 0;">
            <p style="margin:0 0 3px;font-size:14px;font-weight:700;color:#111111;">Date de retrait</p>
            <p style="margin:0;font-size:12px;color:#777777;line-height:1.6;">${dateLabel}</p>
          </td>
        </tr>
      </table>

      <!-- Condition 3 -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="46" valign="top" style="padding:14px 16px 14px 0;">
            <table cellpadding="0" cellspacing="0" border="0" style="background:#111111;border-radius:4px;width:38px;height:38px;">
              <tr><td align="center" valign="middle" style="font-size:18px;width:38px;height:38px;">👥</td></tr>
            </table>
          </td>
          <td valign="top" style="padding:14px 0;">
            <p style="margin:0 0 3px;font-size:14px;font-weight:700;color:#111111;">5 personnes minimum</p>
            <p style="margin:0;font-size:12px;color:#777777;line-height:1.6;">Valable autour d'un repas partagé en groupe d'<strong style="color:#111111;">au moins 5 personnes</strong>.</p>
          </td>
        </tr>
      </table>

      <!-- CTA -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:36px;">
        <tr><td align="center">
          <a href="https://ted-crm.pages.dev/reserver.html" style="display:inline-block;background:#F0A830;color:#111111;font-size:12px;font-weight:700;letter-spacing:3px;text-transform:uppercase;text-decoration:none;padding:18px 48px;border-radius:4px;">Réserver ma table</a>
        </td></tr>
        <tr><td align="center" style="padding-top:14px;">
          <p style="margin:0;font-size:12px;color:#aaaaaa;">ou appelez-nous &nbsp;·&nbsp; <a href="tel:0478906780" style="color:#F0A830;text-decoration:none;font-weight:600;">04 78 90 67 80</a></p>
        </td></tr>
      </table>

      <!-- Signature -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:48px;border-top:1px solid #ece8e0;">
        <tr><td align="center" style="padding-top:32px;">
          <p style="margin:0 0 4px;font-style:italic;font-size:17px;color:#555555;line-height:1.85;">On vous attend avec impatience.</p>
          <p style="margin:0 0 16px;font-size:15px;color:#777777;">À très bientôt,</p>
          <p class="alex" style="margin:0;font-family:'Alex Brush',Georgia,serif;font-size:34px;color:#F0A830;">L'équipe du TED</p>
        </td></tr>
      </table>

    </td></tr>

    <!-- ══ FOOTER ══ -->
    <tr><td bgcolor="#0d0d0d" style="background-color:#0d0d0d;border-radius:0 0 10px 10px;padding:32px 40px;text-align:center;">
      <img src="${logoUrl}" width="44" height="44" alt="Le TED" style="display:block;margin:0 auto 12px;border-radius:50%;opacity:0.7;">
      <p style="margin:0 0 14px;font-size:12px;font-weight:800;letter-spacing:7px;color:#F0A830;text-transform:uppercase;">Le Ted</p>
      <p style="margin:0 0 10px;font-size:11px;color:rgba(255,255,255,0.25);line-height:2;">
        28 Avenue des Frères Montgolfier, 69680 Chassieu<br>
        <a href="tel:0478906780" style="color:rgba(240,168,48,0.5);text-decoration:none;">04 78 90 67 80</a>
        &nbsp;·&nbsp;
        <a href="https://letedchassieu.fr/reservation" style="color:rgba(240,168,48,0.5);text-decoration:none;">leted.fr</a>
      </p>
      <p style="margin:0;font-size:10px;color:rgba(255,255,255,0.12);line-height:1.7;">
        &copy; 2026 Le TED &mdash; Restaurant &amp; Club &mdash; Chassieu, Lyon
      </p>
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
