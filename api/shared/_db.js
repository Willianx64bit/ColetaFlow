import pg from 'pg';
import crypto from 'node:crypto';
const { Pool } = pg;
let pool;
export function db() {
  if (!process.env.DATABASE_URL) throw new Error('DATABASE_URL não configurada');
  if (!pool) pool = new Pool({ connectionString:process.env.DATABASE_URL,max:5,idleTimeoutMillis:10000,connectionTimeoutMillis:8000,ssl:{rejectUnauthorized:true} });
  return pool;
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
  const r=await db().query(`SELECT p.id,p.collection_id,p.role,p.display_name,p.status,c.status AS collection_status
    FROM cf_participants p JOIN cf_collections c ON c.id=p.collection_id
    WHERE p.token_hash=$1 AND p.collection_id=$2 AND p.status='active'`,[hashToken(token),collectionId]);
  return r.rows[0]||null;
}

export async function loggedUser(req) {
 const token=tokenFrom(req); if(!token)return null;
 const r=await db().query(`SELECT u.id,u.email,u.display_name,u.company_name,u.role,u.status,
    EXISTS(SELECT 1 FROM cf_platform_owner o WHERE o.user_id=u.id) AS is_owner
  FROM cf_user_sessions s JOIN cf_users u ON u.id=s.user_id
  WHERE s.token_hash=$1 AND s.expires_at>now() AND u.status='active'`,[hashToken(token)]);
 return r.rows[0]||null;
}
