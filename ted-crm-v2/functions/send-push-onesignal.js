export async function onRequestPost(context) {
  const { title, body } = await context.request.json();

  const res = await fetch('https://api.onesignal.com/notifications', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Key ${context.env.ONESIGNAL_REST_API_KEY}`
    },
    body: JSON.stringify({
      app_id: '87b29550-ffb0-412a-9682-05fdace514fc',
      target_channel: 'push',
      include_aliases: {
        external_id: ['ted-admin']
      },
      headings: { en: title },
      contents: { en: body },
      web_push_topic: 'nouvelle-resa'
    })
  });

  const data = await res.json();
  console.log('OneSignal response:', res.status, JSON.stringify(data));
  return Response.json({ success: res.ok, data });
}
