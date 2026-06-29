export async function onRequest(context) {
  const { request, env } = context;

  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  if (request.method === 'OPTIONS') {
    return new Response(null, { headers });
  }

  const body = await request.json();
  const { client_id, tel, prenom, nom, date, heure, nb_personnes } = body;
  console.log('vapi-create-resa reçu:', JSON.stringify({ client_id, tel, prenom, nom, date, heure, nb_personnes }));

  let clientId = client_id;

  if (!clientId) {
    const resClient = await fetch(
      `${env.REACT_APP_SUPABASE_URL}/rest/v1/clients`,
      {
        method: 'POST',
        headers: {
          'apikey': env.SUPABASE_SERVICE_KEY,
          'Authorization': `Bearer ${env.SUPABASE_SERVICE_KEY}`,
          'Content-Type': 'application/json',
          'Prefer': 'return=representation'
        },
        body: JSON.stringify({ tel, prenom, nom, genre: 'Homme' })
      }
    );
    const newClient = await resClient.json();
    console.log('Création client:', resClient.status, JSON.stringify(newClient));
    clientId = newClient[0]?.id;
  }

  const heureNum = parseInt(heure.split(':')[0]);
  const service = heureNum < 15 ? 'midi' : 'soir';

  const resResa = await fetch(
    `${env.REACT_APP_SUPABASE_URL}/rest/v1/reservations`,
    {
      method: 'POST',
      headers: {
        'apikey': env.SUPABASE_SERVICE_KEY,
        'Authorization': `Bearer ${env.SUPABASE_SERVICE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=representation'
      },
      body: JSON.stringify({
        client_id: clientId,
        date,
        heure,
        service,
        nb_personnes: parseInt(nb_personnes),
        statut: 'attente'
      })
    }
  );

  const resaData = await resResa.json();
  console.log('Réponse Supabase réservation:', resResa.status, JSON.stringify(resaData));

  return new Response(JSON.stringify({
    success: resResa.ok,
    status: resResa.status,
    data: resaData
  }), { headers });
}
