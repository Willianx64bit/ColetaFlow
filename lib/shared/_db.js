import pg from 'pg';
import crypto from 'node:crypto';

const { Client } = pg;
let runtimeEnv = null;

export function configureRuntime(env) {
  runtimeEnv = env;
}

function connectionString() {
  const value = runtimeEnv?.HYPERDRIVE?.connectionString || runtimeEnv?.DATABASE_URL;
  if (!value) throw new Error('Binding HYPERDRIVE não configurado');
  return value;
}

function createClient() {
  return new Client({
    connectionString: connectionString(),
    connectionTimeoutMillis: 8000,
    application_name: 'ColetaFlow'
  });
}

// Hyperdrive owns the upstream connection pool. A short-lived pg Client is
// created per query/request, while explicit transactions keep one client.
export function db() {
  return {
    async query(...args) {
      const client = createClient();
      await client.connect();
      try {
        return await client.query(...args);
      } finally {
        await client.end();
      }
    },
    async connect() {
      const client = createClient();
      await client.connect();
      client.release = () => { void client.end(); };
      return client;
    }
  };
}

export function runtimeSecret(name) {
  return runtimeEnv?.[name] || '';
}

export function send(res,status,body) { res.status(status).setHeader('Cache-Control','no-store').json(body); }
export async function readJson(req) {
  if (req.body && typeof req.body==='object') return req.body;
  let raw=''; for await (const chunk of req) raw+=chunk;
  try { return JSON.parse(raw||'{}'); } catch { return {}; }
}
export function hashToken(token) { return crypto.createHash('sha256').update(token).digest('hex'); }
export function tokenFrom(req) { const h=req.headers.authorization||''; return h.startsWith('Bearer ')?h.slice(7).trim():''; }
export async function participant(req,collectionId) {
  const token=tokenFrom(req); if (!token) return null;
  const user=await loggedUser(req);
  if(user){
    const byAccount=await db().query(`SELECT p.id,p.collection_id,p.role,p.display_name,p.status,c.status AS collection_status
      FROM cf_participants p JOIN cf_collections c ON c.id=p.collection_id
      WHERE p.user_id=$1 AND p.collection_id=$2
        AND (p.status='active' OR (user.is_owner AND p.role='owner' AND p.status='finished'))`,[user.id,collectionId]);
    if(byAccount.rows[0])return byAccount.rows[0];
  }
  const byCollectionToken=await db().query(`SELECT p.id,p.collection_id,p.role,p.display_name,p.status,c.status AS collection_status
    FROM cf_participants p
    JOIN cf_collections c ON c.id=p.collection_id
    LEFT JOIN cf_users u ON u.id=p.user_id
    WHERE p.token_hash=$1 AND p.collection_id=$2 AND p.status='active'
      AND (p.user_id IS NULL OR u.status='active')`,[hashToken(token),collectionId]);
  return byCollectionToken.rows[0]||null;
}

export async function loggedUser(req) {
 const token=tokenFrom(req); if(!token)return null;
 const r=await db().query(`SELECT u.id,u.email,u.display_name,u.company_name,u.role,u.status,
    EXISTS(SELECT 1 FROM cf_platform_owner o WHERE o.user_id=u.id) AS is_owner
  FROM cf_user_sessions s JOIN cf_users u ON u.id=s.user_id
  WHERE s.token_hash=$1 AND s.expires_at>now() AND u.status='active'`,[hashToken(token)]);
 return r.rows[0]||null;
}
