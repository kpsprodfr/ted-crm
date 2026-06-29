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

  const SUPA_URL = env.REACT_APP_SUPABASE_URL;
  const SUPA_KEY = env.SUPABASE_SERVICE_KEY || env.REACT_APP_SUPABASE_KEY;

  console.log('SUPA_URL:', SUPA_URL ? 'OK' : 'MANQUANT');
  console.log('SUPA_KEY:', SUPA_KEY ? 'OK' : 'MANQUANT');
  console.log('Données reçues:', JSON.stringify(body));

  let clientId = client_id;

  if (!clientId) {
    const resClient = await fetch(`${SUPA_URL}/rest/v1/clients`, {
      method: 'POST',
      headers: {
        'apikey': SUPA_KEY,
        'Authorization': `Bearer ${SUPA_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=representation'
      },
      body: JSON.stringify({ tel, prenom, nom, genre: 'Homme' })
    });
    const newClient = await resClient.json();
    console.log('Client créé:', resClient.status, JSON.stringify(newClient));
    clientId = newClient[0]?.id;
  }

  if (!clientId) {
    return new Response(JSON.stringify({ success: false, error: 'client_id manquant' }), { headers });
  }

  const heureNum = parseInt((heure || '19:00').split(':')[0]);
  const service = heureNum < 15 ? 'midi' : 'soir';

  const resResa = await fetch(`${SUPA_URL}/rest/v1/reservations`, {
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
      nb_personnes: parseInt(nb_personnes) || 2,
      statut: 'attente',
      source: 'telephone'
    })
  });

  const resaData = await resResa.json();
  console.log('Réservation créée:', resResa.status, JSON.stringify(resaData));

  return new Response(JSON.stringify({
    success: resResa.ok,
    status: resResa.status,
    data: resaData
  }), { headers });
}
