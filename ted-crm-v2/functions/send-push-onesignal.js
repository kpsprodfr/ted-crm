export async function onRequestPost(context) {
  const { title, body } = await context.request.json();
  const appId = '87b29550-ffb0-412a-9682-05fdace514fc';
  const apiKey = context.env.ONESIGNAL_REST_API_KEY;

  // 1. Récupère tous les abonnements
  const subsRes = await fetch(`https://api.onesignal.com/apps/${appId}/subscriptions?limit=50`, {
    headers: { 'Authorization': `Key ${apiKey}` }
  });
  const subsData = await subsRes.json();
  const subscriptions = subsData.subscriptions || [];

  // 2. Garde uniquement le plus récent par appareil (basé sur last_active)
  const deviceMap = {};
  for (const sub of subscriptions) {
    const device = sub.device_model || sub.device_type || 'unknown';
    if (!deviceMap[device] || sub.last_active > deviceMap[device].last_active) {
      deviceMap[device] = sub;
    }
  }
  const latestIds = Object.values(deviceMap).map(s => s.id);

  // 3. Supprime les anciens doublons
  for (const sub of subscriptions) {
    if (!latestIds.includes(sub.id)) {
      await fetch(`https://api.onesignal.com/apps/${appId}/subscriptions/${sub.id}`, {
        method: 'DELETE',
        headers: { 'Authorization': `Key ${apiKey}` }
      });
    }
  }

  // 4. Envoie la notification
  const res = await fetch('https://api.onesignal.com/notifications', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Key ${apiKey}`
    },
    body: JSON.stringify({
      app_id: appId,
      included_segments: ['All'],
      headings: { en: title },
      contents: { en: body },
      collapse_id: 'nouvelle-resa'
    })
  });

  const data = await res.json();
  return Response.json({ success: res.ok, data });
}
