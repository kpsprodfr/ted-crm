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
  const { tel } = body;
  const SUPA_URL = env.REACT_APP_SUPABASE_URL;
  const SUPA_KEY = env.SUPABASE_SERVICE_KEY || env.REACT_APP_SUPABASE_KEY;

  const chiffres = tel.replace(/\D/g, '');

  const res = await fetch(
    `${SUPA_URL}/rest/v1/clients?select=id,prenom,nom,tel,genre,entreprise&limit=1000`,
    {
      headers: {
        'apikey': SUPA_KEY,
        'Authorization': `Bearer ${SUPA_KEY}`
      }
    }
  );

  const clients = await res.json();

  if (!Array.isArray(clients)) {
    return new Response(JSON.stringify({ found: false, error: 'Supabase error', raw: clients }), { headers });
  }

  const client = clients.find(c => {
    if (!c.tel) return false;
    const clientChiffres = c.tel.replace(/\D/g, '');
    return clientChiffres.slice(-9) === chiffres.slice(-9);
  });

  if (client) {
    return new Response(JSON.stringify({
      found: true,
      id: client.id,
      prenom: client.prenom || client.entreprise,
      nom: client.nom,
      genre: client.genre
    }), { headers });
  }

  return new Response(JSON.stringify({ found: false, total_clients: clients.length }), { headers });
}
