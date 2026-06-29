export async function onRequest(context) {
  const { env } = context;
  const SUPA_URL = env.REACT_APP_SUPABASE_URL;
  const serviceKey = env.SUPABASE_SERVICE_KEY;
  const anonKey = env.REACT_APP_SUPABASE_KEY;

  // Test avec service key
  let serviceCount = 'N/A';
  if (serviceKey) {
    const r = await fetch(`${SUPA_URL}/rest/v1/reservations?statut=eq.attente&select=id`, {
      headers: { 'apikey': serviceKey, 'Authorization': `Bearer ${serviceKey}` }
    });
    const d = await r.json();
    serviceCount = Array.isArray(d) ? d.length : JSON.stringify(d);
  }

  // Test avec anon key
  let anonCount = 'N/A';
  if (anonKey) {
    const r = await fetch(`${SUPA_URL}/rest/v1/reservations?statut=eq.attente&select=id`, {
      headers: { 'apikey': anonKey, 'Authorization': `Bearer ${anonKey}` }
    });
    const d = await r.json();
    anonCount = Array.isArray(d) ? d.length : JSON.stringify(d);
  }

  return new Response(JSON.stringify({
    serviceKeyPresent: !!serviceKey,
    anonKeyPresent: !!anonKey,
    serviceCount_attente: serviceCount,
    anonCount_attente: anonCount
  }), { headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' } });
}
