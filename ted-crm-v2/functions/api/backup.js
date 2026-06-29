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

  // Protection par token simple
  const authHeader = request.headers.get('Authorization') || '';
  const token = authHeader.replace('Bearer ', '');
  if (token !== env.BACKUP_SECRET) {
    return new Response(JSON.stringify({ error: 'Non autorisé' }), { status: 401, headers });
  }

  const SUPA_URL = env.REACT_APP_SUPABASE_URL;
  const SUPA_KEY = env.SUPABASE_SERVICE_KEY || env.REACT_APP_SUPABASE_KEY;

  const supaHeaders = {
    'apikey': SUPA_KEY,
    'Authorization': `Bearer ${SUPA_KEY}`
  };

  const tables = ['clients', 'reservations', 'sms_envoyes', 'emails_envoyes'];

  const results = {};
  for (const table of tables) {
    try {
      const res = await fetch(`${SUPA_URL}/rest/v1/${table}?select=*&limit=10000`, {
        headers: supaHeaders
      });
      results[table] = await res.json();
    } catch (e) {
      results[table] = { error: e.message };
    }
  }

  const date = new Date().toISOString().split('T')[0];
  const backup = {
    date,
    exported_at: new Date().toISOString(),
    tables: results,
    counts: Object.fromEntries(
      tables.map(t => [t, Array.isArray(results[t]) ? results[t].length : 0])
    )
  };

  return new Response(JSON.stringify(backup, null, 2), {
    headers: {
      ...headers,
      'Content-Disposition': `attachment; filename="backup-ted-${date}.json"`
    }
  });
}
