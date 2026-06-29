export async function onRequest(context) {
  const { request, env } = context;
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Content-Type': 'application/json'
  };

  if (request.method === 'OPTIONS') {
    return new Response(null, { headers });
  }

  // Protection par token si BACKUP_SECRET est défini
  const authHeader = request.headers.get('Authorization') || '';
  const token = authHeader.replace('Bearer ', '');
  if (env.BACKUP_SECRET && token !== env.BACKUP_SECRET) {
    return new Response(JSON.stringify({ error: 'Non autorisé' }), { status: 401, headers });
  }

  const SUPA_URL = env.REACT_APP_SUPABASE_URL;
  const SUPA_KEY = env.SUPABASE_SERVICE_KEY || env.REACT_APP_SUPABASE_KEY;
  const supaHeaders = {
    'apikey': SUPA_KEY,
    'Authorization': `Bearer ${SUPA_KEY}`,
    'Prefer': 'count=exact'
  };

  // Compte les réservations en attente (head request, juste le count)
  const resAttente = await fetch(
    `${SUPA_URL}/rest/v1/reservations?statut=eq.attente&select=id`,
    { method: 'HEAD', headers: supaHeaders }
  );
  const countAttente = parseInt((resAttente.headers.get('content-range') || '0/0').split('/')[1]) || 0;

  // Liste les 10 dernières en attente avec le détail
  const resList = await fetch(
    `${SUPA_URL}/rest/v1/reservations?statut=eq.attente&select=id,date,heure,nb_personnes,service,source,created_at,clients(prenom,nom,tel)&order=created_at.desc&limit=10`,
    { headers: { 'apikey': SUPA_KEY, 'Authorization': `Bearer ${SUPA_KEY}` } }
  );
  const dernieres = await resList.json();

  return new Response(JSON.stringify({
    attente_total: countAttente,
    dernieres_attente: dernieres
  }, null, 2), { headers });
}
