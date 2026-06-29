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

  const res = await fetch(
    `${env.REACT_APP_SUPABASE_URL}/rest/v1/clients?tel=eq.${tel}&select=*`,
    {
      headers: {
        'apikey': env.REACT_APP_SUPABASE_KEY,
        'Authorization': `Bearer ${env.REACT_APP_SUPABASE_KEY}`
      }
    }
  );

  const clients = await res.json();

  if (clients && clients.length > 0) {
    const client = clients[0];
    return new Response(JSON.stringify({
      found: true,
      id: client.id,
      prenom: client.prenom || client.entreprise,
      nom: client.nom,
      genre: client.genre
    }), { headers });
  }

  return new Response(JSON.stringify({ found: false }), { headers });
}
