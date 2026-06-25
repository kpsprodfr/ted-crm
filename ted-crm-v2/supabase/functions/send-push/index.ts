import { serve } from "https://deno.land/std@0.168.0/http/server.ts"

serve(async (req) => {
  const { record } = await req.json()

  const title = '📅 Nouvelle réservation !'
  const body = `${record.nb_personnes} pers. · ${record.heure} · ${record.service === 'midi' ? 'Déjeuner' : 'Dîner'}`

  await fetch('https://api.onesignal.com/notifications', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Key ${Deno.env.get('ONESIGNAL_REST_API_KEY')}`
    },
    body: JSON.stringify({
      app_id: '87b29550-ffb0-412a-9682-05fdace514fc',
      included_segments: ['All'],
      headings: { en: title },
      contents: { en: body }
    })
  })

  return new Response(JSON.stringify({ success: true }), {
    headers: { 'Content-Type': 'application/json' }
  })
})
