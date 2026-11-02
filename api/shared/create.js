import { randomBytes,createHash } from 'node:crypto';
import { db,readJson,send } from './_db.js';
export default async function handler(req,res) {
 if(req.method!=='POST') return send(res,405,{error:'Método não permitido'});
 try {
  const body=await readJson(req), managerName=String(body.managerName||'').trim().slice(0,60), title=String(body.title||'Coleta compartilhada').trim().slice(0,120);
  if(!managerName) return send(res,400,{error:'Informe o nome do gerente'});
  const code=randomBytes(4).toString('hex').toUpperCase(), token=randomBytes(32).toString('base64url'), hash=createHash('sha256').update(token).digest('hex');
  const client=await db().connect();
  try {
   await client.query('BEGIN');
   const c=await client.query("INSERT INTO cf_collections(access_code,title) VALUES($1,$2) RETURNING id,access_code,title,status,created_at",[code,title]);
   const p=await client.query("INSERT INTO cf_participants(collection_id,display_name,token_hash,role) VALUES($1,$2,$3,'manager') RETURNING id",[c.rows[0].id,managerName,hash]);
   await client.query('COMMIT');
   return send(res,201,{collection:c.rows[0],participantId:p.rows[0].id,token});
  } catch(e){await client.query('ROLLBACK');throw e;} finally{client.release();}
 } catch(e){console.error('shared create',e);return send(res,500,{error:'Não foi possível criar a coleta'});}
}
