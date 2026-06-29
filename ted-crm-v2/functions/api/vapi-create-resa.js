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

  console.log('ENV check:', {
    supabaseUrl: env.REACT_APP_SUPABASE_URL ? 'OK' : 'MANQUANT',
    serviceKey: SUPA_KEY ? 'OK' : 'MANQUANT',
    publishableKey: env.REACT_APP_SUPABASE_KEY ? 'OK' : 'MANQUANT'
  });

  const SUPA_KEY = SUPA_KEY || env.REACT_APP_SUPABASE_KEY;

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
          'apikey': SUPA_KEY,
          'Authorization': `Bearer ${SUPA_KEY}`,
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
        'apikey': SUPA_KEY,
        'Authorization': `Bearer ${SUPA_KEY}`,
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
