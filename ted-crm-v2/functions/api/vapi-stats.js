export async function onRequest(context) {
  const { env } = context;
  const SUPA_URL = env.REACT_APP_SUPABASE_URL;
  const SUPA_KEY = env.SUPABASE_SERVICE_KEY;
  const res = await fetch(`${SUPA_URL}/rest/v1/reservations?statut=eq.attente&select=id,date,heure,nb_personnes,clients(prenom,nom)&order=created_at.desc`, {
    headers: { 'apikey': SUPA_KEY, 'Authorization': `Bearer ${SUPA_KEY}` }
  });
  const data = await res.json();
  return new Response(JSON.stringify({ attente_total: Array.isArray(data) ? data.length : 0, data }), {
    headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
  });
}
