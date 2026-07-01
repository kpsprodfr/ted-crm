// Endpoint legacy (plus aucun appelant dans le code) — conservé mais verrouillé :
// JWT CRM requis, origine restreinte, rate limit, échappement HTML.
import { guard, secureJson, verifyUser, escapeHtml, isValidEmail, isValidTelFR } from '../_utils.js';

export async function onRequestPost(context) {
  const { env, request } = context;

  const blocked = await guard(env, request, { limit: 10, bucket: 'roue-notify' });
  if (blocked) return blocked;

  const user = await verifyUser(env, request);
  if (!user) return secureJson({ error: 'Non autorisé' }, { status: 401 });

  const apiKey = env.BREVO_API_KEY;
  const SUPA_URL = env.SUPABASE_URL || 'https://mwpfaytccypvdrgapptk.supabase.co';
  const SUPA_KEY = env.SUPABASE_ANON_KEY || env.SUPABASE_SERVICE_ROLE_KEY;

  let body;
  try { body = await request.json(); } catch { return secureJson({ error: 'JSON invalide' }, { status: 400 }); }

  const { tel, email, prenom, recompense_nom, recompense_emoji, conditions, date_debut_validite } = body;
  if (!tel || !email) return secureJson({ error: 'tel et email requis' }, { status: 400 });
  if (!isValidEmail(email)) return secureJson({ error: 'email invalide' }, { status: 400 });
  if (!isValidTelFR(tel)) return secureJson({ error: 'tel invalide' }, { status: 400 });

  const prenomStr = escapeHtml(prenom ? String(prenom).trim().slice(0, 50) : '');
  const recompenseSafe = escapeHtml(recompense_nom || '');
  const emojiSafe = escapeHtml(recompense_emoji || '');
  const conditionsSafe = escapeHtml(conditions || '');
  const dateStr = date_debut_validite
    ? new Date(date_debut_validite).toLocaleDateString('fr-FR', { day: '2-digit', month: 'long', year: 'numeric' })
    : "dès maintenant";

  const results = {};

  // ── SMS ─────────────────────────────────────────────────────────────────────
  if (apiKey) {
    const smsContent = [
      `🎉 Félicitations${prenomStr ? ' ' + prenomStr : ''} !`,
      `Vous avez gagné ${recompense_emoji || ''} ${recompense_nom} au TED !`,
      `Valable ${dateStr}.`,
      conditions ? conditions : '',
      `À présenter en caisse.`
    ].filter(Boolean).join(' ');

    const numeroNettoye = tel.replace(/[\s.\-()]/g, '').replace(/^0/, '+33');
    try {
      const smsRes = await fetch('https://api.brevo.com/v3/transactionalSMS/sms', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': apiKey },
        body: JSON.stringify({
          sender: 'LE TED',
          recipient: numeroNettoye,
          content: smsContent,
          type: 'transactional',
          unicodeEnabled: true,
        }),
      });
      results.sms = { ok: smsRes.ok, status: smsRes.status };
    } catch (e) { results.sms = { ok: false, error: e.message }; }
  }

  // ── Email ───────────────────────────────────────────────────────────────────
  if (apiKey) {
    const htmlEmail = `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:'Segoe UI',Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f5f5;padding:32px 16px;">
    <tr><td align="center">
      <table width="100%" style="max-width:520px;background:#111;border-radius:20px;overflow:hidden;">
        <!-- Header -->
        <tr><td style="background:#111;padding:28px 32px;text-align:center;border-bottom:3px solid #E8C547;">
          <div style="font-size:28px;font-weight:900;color:#E8C547;letter-spacing:3px;">LE TED</div>
          <div style="font-size:11px;color:#666;letter-spacing:2px;margin-top:4px;text-transform:uppercase;">Restaurant &amp; Club · Chassieu</div>
        </td></tr>
        <!-- Body -->
        <tr><td style="padding:32px;">
          <div style="text-align:center;font-size:48px;margin-bottom:16px;">${emojiSafe || '🎉'}</div>
          <h1 style="color:#fff;font-size:22px;font-weight:800;text-align:center;margin:0 0 8px;">Félicitations${prenomStr ? ' ' + prenomStr : ''}&nbsp;!</h1>
          <p style="color:#999;font-size:14px;text-align:center;margin:0 0 24px;">Vous avez participé à la roue des cadeaux du TED.</p>

          <div style="background:#1a1a1a;border:2px solid #E8C547;border-radius:14px;padding:24px;text-align:center;margin-bottom:24px;">
            <div style="font-size:12px;font-weight:700;color:#E8C547;letter-spacing:2px;text-transform:uppercase;margin-bottom:10px;">Votre récompense</div>
            <div style="font-size:26px;font-weight:900;color:#fff;">${emojiSafe} ${recompenseSafe}</div>
            ${conditionsSafe ? `<div style="font-size:13px;color:#888;margin-top:10px;">${conditionsSafe}</div>` : ''}
            <div style="font-size:13px;color:#E8C547;font-weight:700;margin-top:12px;">Valable ${dateStr}</div>
          </div>

          <div style="background:#1c1c1c;border-radius:10px;padding:16px;margin-bottom:24px;">
            <div style="font-size:12px;color:#666;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Comment en profiter ?</div>
            <div style="font-size:14px;color:#ccc;line-height:1.6;">
              Présentez cet email ou le SMS reçu <strong style="color:#fff;">à votre serveur ou en caisse</strong> lors de votre prochain passage au TED.
            </div>
          </div>

          <div style="text-align:center;padding-top:8px;border-top:1px solid #222;">
            <div style="font-size:12px;color:#555;">28 Av. des Frères Montgolfier · 69680 Chassieu</div>
            <div style="font-size:12px;color:#555;margin-top:4px;">04 72 02 20 20</div>
          </div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;

    try {
      const emailRes = await fetch('https://api.brevo.com/v3/smtp/email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': apiKey },
        body: JSON.stringify({
          sender: { name: 'Le TED', email: 'com.astegal@gmail.com' },
          to: [{ email, name: prenomStr || 'Client' }],
          subject: `🎉 ${prenomStr ? prenomStr + ', v' : 'V'}ous avez gagné ${recompense_emoji || ''} ${String(recompense_nom).slice(0,80)} !`,
          htmlContent: htmlEmail,
        }),
      });
      results.email = { ok: emailRes.ok, status: emailRes.status };
    } catch (e) { results.email = { ok: false, error: e.message }; }
  }

  return secureJson({ success: true, results });
}
