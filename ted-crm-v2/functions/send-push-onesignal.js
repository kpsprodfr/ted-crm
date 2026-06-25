export async function onRequestPost(context) {
  console.log('SEND-PUSH APPELÉ', new Date().toISOString());
  const { title, body } = await context.request.json();

  const res = await fetch('https://api.onesignal.com/notifications', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Key ${context.env.ONESIGNAL_REST_API_KEY}`
    },
    body: JSON.stringify({
      app_id: '87b29550-ffb0-412a-9682-05fdace514fc',
      included_segments: ['All'],
      headings: { en: title },
      contents: { en: body }
    })
  });

  const data = await res.json();
  return Response.json({ success: res.ok, data });
}
