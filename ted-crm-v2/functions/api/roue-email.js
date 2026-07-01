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

  const defaultObjet = '🎉 {prenom}, vous avez gagné {emoji} {recompense} au Grand Jeux du TED !';
  let objet = p[objetCle] || defaultObjet;
  const messagePerso = p['roue_email_message'] || '';

  // Date
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

  const htmlContent = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://fonts.googleapis.com/css2?family=Alex+Brush&display=swap" rel="stylesheet">
<style>
  @keyframes shimmer {
    0%   { background-position: -400% center; }
    100% { background-position:  400% center; }
  }
  @keyframes pulse-ring {
    0%, 100% { opacity: 0.4; transform: scale(1); }
    50%       { opacity: 0.8; transform: scale(1.04); }
  }
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(24px); }
    to   { opacity: 1; transform: translateY(0); }
  }
  @keyframes confetti-drop {
    0%   { transform: translateY(-10px) rotate(0deg); opacity: 1; }
    100% { transform: translateY(60px)  rotate(180deg); opacity: 0; }
  }
  body { margin: 0; background: #1a1a1a; font-family: -apple-system, 'SF Pro Display', 'Helvetica Neue', sans-serif; }
  * { box-sizing: border-box; }
</style>
</head>
<body>
<div style="display:none;max-height:0;overflow:hidden;font-size:1px;color:#1a1a1a;">${preheader}&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;</div>

<div style="min-height: 100vh; background: #181818; padding: 40px 16px; display: flex; flex-direction: column; align-items: center;">

  <div style="width: 100%; max-width: 600px; animation: fadeUp 0.6s ease both;">

    <!-- HEADER -->
    <div style="background: #111111; border-radius: 10px 10px 0 0; padding: 44px 40px 36px; text-align: center; position: relative; overflow: hidden;">
      <div style="position: absolute; top: 0; left: 50%; transform: translateX(-50%); width: 320px; height: 200px; background: radial-gradient(ellipse at 50% 0%, rgba(240,168,48,0.22) 0%, transparent 70%); pointer-events: none;"></div>
      <img src="${logoUrl}" alt="Le TED" width="72" height="72" style="object-fit: contain; display: block; margin: 0 auto 18px; position: relative;" />
      <div style="font-weight: 800; font-size: 26px; letter-spacing: 10px; text-transform: uppercase; background: linear-gradient(90deg, #c47e10 0%, #F0A830 30%, #ffd278 50%, #F0A830 70%, #c47e10 100%); background-size: 300% auto; -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; animation: shimmer 6s linear infinite; margin-bottom: 6px;">Le Ted</div>
      <div style="font-size: 10px; letter-spacing: 4px; color: rgba(255,255,255,0.25); text-transform: uppercase; margin-top: 6px;">Restaurant &amp; Club · Chassieu</div>
    </div>

    <!-- Gold bar -->
    <div style="height: 3px; background: linear-gradient(90deg, #c47e10, #F0A830, #ffd278, #F0A830, #c47e10); background-size: 300% auto; animation: shimmer 4s linear infinite;"></div>

    <!-- WINNER HERO -->
    <div style="background: #111111; padding: 48px 40px 44px; text-align: center; border-top: none;">

      <div style="position: relative; display: flex; justify-content: center; margin-bottom: 12px;">
        <div style="position: absolute; top: -8px; left: 60px; width: 6px; height: 6px; background: #F0A830; transform: rotate(20deg); animation: confetti-drop 2.4s ease-in infinite;"></div>
        <div style="position: absolute; top: -4px; left: 90px; width: 4px; height: 4px; background: #fff; border-radius: 50%; animation: confetti-drop 2.1s ease-in 0.3s infinite;"></div>
        <div style="position: absolute; top: -12px; right: 60px; width: 5px; height: 5px; background: #F0A830; transform: rotate(-15deg); animation: confetti-drop 2.6s ease-in 0.6s infinite;"></div>
        <div style="position: absolute; top: -6px; right: 88px; width: 3px; height: 8px; background: rgba(240,168,48,0.5); transform: rotate(30deg); animation: confetti-drop 2.3s ease-in 1s infinite;"></div>
        <div style="font-weight: 800; font-size: 11px; letter-spacing: 5px; color: #F0A830; text-transform: uppercase; opacity: 0.9;">✦ Vous avez gagné ✦</div>
      </div>

      <h1 style="font-size: 38px; font-weight: 700; color: #ffffff; margin: 16px 0 8px; line-height: 1.15;">${to_prenom || ''} ${to_nom || ''}</h1>
      <p style="font-size: 13px; color: rgba(255,255,255,0.45); margin: 0 0 44px; letter-spacing: 0.5px;">vous repart avec une récompense exclusive</p>

      <!-- BON GAGNANT -->
      <div style="position: relative; margin-bottom: 44px;">
        <div style="position: absolute; inset: -8px; border-radius: 20px; background: radial-gradient(ellipse at 50% 50%, rgba(240,168,48,0.15) 0%, transparent 70%); animation: pulse-ring 3s ease-in-out infinite; pointer-events: none;"></div>

        <div style="border-radius: 14px; overflow: hidden; border: 1px solid rgba(240,168,48,0.35); position: relative;">
          <div style="height: 4px; background: linear-gradient(90deg, #c47e10, #F0A830, #ffd278, #F0A830, #c47e10); background-size: 300% auto; animation: shimmer 4s linear infinite;"></div>

          <div style="background: #1a1300; padding: 14px 24px; display: flex; align-items: center; justify-content: space-between;">
            <div style="display: flex; align-items: center; gap: 8px;">
              <img src="${logoUrl}" alt="TED" width="22" height="22" style="object-fit: contain; opacity: 0.85;" />
              <span style="font-size: 9px; font-weight: 700; letter-spacing: 3.5px; color: #F0A830; text-transform: uppercase;">Bon Gagnant</span>
            </div>
            <span style="font-family: 'Courier New', monospace; font-size: 10px; color: rgba(240,168,48,0.3); letter-spacing: 1px;">${serialCode}</span>
          </div>

          <div style="background: #111; border-top: 1px dashed rgba(240,168,48,0.2); border-bottom: 1px dashed rgba(240,168,48,0.2); height: 1px; margin: 0;"></div>

          <div style="background: linear-gradient(160deg, #161000 0%, #111111 100%); padding: 40px 28px 32px; text-align: center;">
            <div style="font-size: 9px; letter-spacing: 7px; color: rgba(240,168,48,0.4); margin-bottom: 20px;">✦ &nbsp; ✦ &nbsp; ✦</div>
            <div style="font-size: 64px; line-height: 1; margin-bottom: 20px; filter: drop-shadow(0 4px 20px rgba(240,168,48,0.35));">${emoji || '🎁'}</div>
            <div style="font-size: 9px; font-weight: 600; letter-spacing: 4px; color: rgba(240,168,48,0.5); text-transform: uppercase; margin-bottom: 12px;">Votre récompense</div>
            <div style="font-size: 26px; font-weight: 700; color: #ffffff; line-height: 1.2; margin-bottom: 28px; letter-spacing: -0.3px;">${recompense || ''}</div>

            <div style="display: inline-flex; align-items: center; gap: 8px; background: rgba(240,168,48,0.08); border: 1px solid rgba(240,168,48,0.2); border-radius: 40px; padding: 8px 20px; margin-bottom: 20px;">
              <div style="width: 5px; height: 5px; background: #F0A830; border-radius: 50%;"></div>
              <span style="font-size: 14px; color: rgba(255,255,255,0.5);">Valable</span>
              <span style="font-size: 14px; font-weight: 600; color: #F0A830;">${dateLabel}</span>
            </div>

            <div style="display: flex; align-items: center; justify-content: center; gap: 5px; margin-bottom: 0;">
              <div style="height: 1px; flex: 1; background: linear-gradient(to right, transparent, rgba(240,168,48,0.2));"></div>
              <div style="font-size: 8px; letter-spacing: 5px; color: rgba(240,168,48,0.25);">· · · · ·</div>
              <div style="height: 1px; flex: 1; background: linear-gradient(to left, transparent, rgba(240,168,48,0.2));"></div>
            </div>
          </div>

          <div style="position: relative; display: flex; align-items: center; background: #111;">
            <div style="width: 16px; height: 16px; background: #111111; border-radius: 50%; flex-shrink: 0; margin-left: -8px; border: 1px solid rgba(240,168,48,0.25);"></div>
            <div style="flex: 1; border-top: 1.5px dashed rgba(240,168,48,0.25);"></div>
            <div style="width: 16px; height: 16px; background: #111111; border-radius: 50%; flex-shrink: 0; margin-right: -8px; border: 1px solid rgba(240,168,48,0.25);"></div>
          </div>

          <div style="background: #0e0e0e; padding: 14px 28px; text-align: center;">
            <a href="#conditions" style="font-size: 12px; color: rgba(255,255,255,0.4); text-decoration: underline; letter-spacing: 0.5px;">Lire les conditions de retrait en dessous ↓</a>
          </div>

          <div style="height: 3px; background: linear-gradient(90deg, #c47e10, #F0A830, #ffd278, #F0A830, #c47e10); background-size: 300% auto; animation: shimmer 4s linear infinite; opacity: 0.6;"></div>
        </div>
      </div>

      ${hasMessage ? `
      <div style="margin-bottom: 44px; padding: 20px 24px; border-left: 2px solid rgba(240,168,48,0.4); text-align: left; background: rgba(255,255,255,0.03); border-radius: 0 6px 6px 0;">
        <p style="font-style: italic; font-size: 17px; color: rgba(255,255,255,0.6); margin: 0 0 8px; line-height: 1.7;">&laquo;&nbsp;${messagePerso.replace(/\n/g, '<br>')}&nbsp;&raquo;</p>
        <p style="font-size: 10px; letter-spacing: 2.5px; color: rgba(240,168,48,0.5); text-transform: uppercase; margin: 0;">L'équipe du TED</p>
      </div>` : ''}
    </div>

    <!-- CONDITIONS -->
    <div id="conditions" style="background: #FDFAF5; padding: 44px 40px 40px;">
      <p style="font-size: 22px; font-weight: 700; letter-spacing: 0.5px; color: #777; text-transform: uppercase; margin: 0 0 24px;">Conditions de retrait</p>

      <div style="display: flex; flex-direction: column; gap: 0;">
        <div style="display: flex; align-items: flex-start; gap: 16px; padding: 16px 0; border-bottom: 1px solid #f0ece4;">
          <div style="width: 38px; height: 38px; background: #111; border-radius: 4px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; font-size: 17px;">📋</div>
          <div style="padding-top: 4px;">
            <div style="font-size: 15px; font-weight: 700; color: #111; margin-bottom: 3px;">Présentation obligatoire</div>
            <div style="font-size: 13px; color: #777; line-height: 1.6;">Présentez <strong style="color: #111; font-weight: 700;">cet email</strong> à votre arrivée au restaurant.</div>
          </div>
        </div>

        <div style="display: flex; align-items: flex-start; gap: 16px; padding: 16px 0; border-bottom: 1px solid #f0ece4;">
          <div style="width: 38px; height: 38px; background: #111; border-radius: 4px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; font-size: 17px;">📅</div>
          <div style="padding-top: 4px;">
            <div style="font-size: 15px; font-weight: 700; color: #111; margin-bottom: 3px;">Date de retrait</div>
            <div style="font-size: 13px; color: #777; line-height: 1.6;">${dateLabel}</div>
          </div>
        </div>

        <div style="display: flex; align-items: flex-start; gap: 16px; padding: 16px 0;">
          <div style="width: 38px; height: 38px; background: #111; border-radius: 4px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; font-size: 17px;">👥</div>
          <div style="padding-top: 4px;">
            <div style="font-size: 15px; font-weight: 700; color: #111; margin-bottom: 3px;">5 personnes minimum</div>
            <div style="font-size: 13px; color: #777; line-height: 1.6;">Valable autour d'un repas partagé en groupe d'<strong style="color: #111; font-weight: 700;">au moins 5 personnes</strong>.</div>
          </div>
        </div>
      </div>

      <div style="margin-top: 40px; text-align: center;">
        <a href="https://letedchassieu.fr/reservation" style="display: inline-block; background: #F0A830; color: #111111; font-size: 12px; font-weight: 700; letter-spacing: 3px; text-transform: uppercase; text-decoration: none; padding: 18px 52px; border-radius: 4px;">Réserver ma table</a>
        <p style="font-size: 12px; color: #bbb; margin: 14px 0 0;">ou appelez-nous · <a href="tel:0478906780" style="color: #F0A830; text-decoration: none; font-weight: 600;">04 78 90 67 80</a></p>
      </div>

      <div style="margin-top: 52px; padding-top: 36px; border-top: 1px solid #ece8e0; text-align: center;">
        <p style="font-style: italic; font-size: 18px; color: #555; line-height: 1.85; margin: 0 0 4px;">On vous attend avec impatience.</p>
        <p style="font-size: 16px; color: #777; margin: 0 0 18px;">À très bientôt,</p>
        <div style="font-family: 'Alex Brush', cursive; font-size: 36px; color: #F0A830;">L'équipe du TED</div>
      </div>
    </div>

    <!-- FOOTER -->
    <div style="background: #0d0d0d; border-radius: 0 0 10px 10px; padding: 36px 40px; text-align: center;">
      <img src="${logoUrl}" alt="Le TED" width="44" height="44" style="object-fit: contain; opacity: 0.7; display: block; margin: 0 auto 12px;" />
      <div style="font-weight: 800; font-size: 13px; letter-spacing: 7px; color: #F0A830; margin-bottom: 16px; text-transform: uppercase;">Le Ted</div>
      <p style="font-size: 11px; color: rgba(255,255,255,0.25); line-height: 2; margin: 0 0 10px;">
        28 Avenue des Frères Montgolfier, 69680 Chassieu<br/>
        <a href="tel:0478906780" style="color: rgba(240,168,48,0.5); text-decoration: none;">04 78 90 67 80</a>
        &nbsp;·&nbsp;
        <a href="https://letedchassieu.fr/reservation" style="color: rgba(240,168,48,0.5); text-decoration: none;">leted.fr</a>
      </p>
      <p style="font-size: 10px; color: rgba(255,255,255,0.12); margin: 0; line-height: 1.7;">
        © 2026 Le TED — Restaurant &amp; Club — Chassieu, Lyon
      </p>
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
