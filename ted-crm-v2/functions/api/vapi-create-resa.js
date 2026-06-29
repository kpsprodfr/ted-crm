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

  console.log('Données reçues:', JSON.stringify({ client_id, tel, prenom, nom, date, heure, nb_personnes }));

  let clientId = client_id;

  if (!clientId && tel) {
    const resSearch = await fetch(
      `${SUPA_URL}/rest/v1/clients?select=id,prenom,nom,tel&limit=1000`,
      { headers: { 'apikey': SUPA_KEY, 'Authorization': `Bearer ${SUPA_KEY}` } }
    );
    const allClients = await resSearch.json();
    const telChiffres = tel.replace(/\D/g, '').slice(-9);
    const found = Array.isArray(allClients) && allClients.find(c => c.tel && c.tel.replace(/\D/g, '').slice(-9) === telChiffres);

    if (found) {
      clientId = found.id;
      console.log('Client existant trouvé:', clientId);
    } else {
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
      console.log('Nouveau client créé:', resClient.status, JSON.stringify(newClient));
      clientId = Array.isArray(newClient) ? newClient[0]?.id : newClient?.id;
    }
  }

  if (!clientId) {
    console.log('ERREUR: pas de client_id');
    return new Response(JSON.stringify({ success: false, error: 'impossible de trouver ou créer le client' }), { headers });
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
    result: resResa.ok ? "Réservation enregistrée avec succès." : "Erreur lors de l'enregistrement."
  }), { headers });
}
