import { randomBytes,createHash } from 'node:crypto';
import { db,readJson,send,loggedUser } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const user=await loggedUser(req);if(!user)return send(res,401,{error:'Entre na sua conta para criar uma coleta'});
  if(!user.is_owner)return send(res,403,{error:'Somente o dono pode criar coletas'});
  const body=await readJson(req),title=String(body.title||'Coleta compartilhada').trim().slice(0,120);
  const code=randomBytes(4).toString('hex').toUpperCase(),token=randomBytes(32).toString('base64url'),hash=createHash('sha256').update(token).digest('hex');
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const c=await client.query("INSERT INTO cf_collections(access_code,title) VALUES($1,$2) RETURNING id,access_code,title,status,created_at",[code,title]);
   const p=await client.query("INSERT INTO cf_participants(collection_id,display_name,token_hash,role,user_id) VALUES($1,$2,$3,'manager',$4) RETURNING id",[c.rows[0].id,user.display_name,hash,user.id]);
   await client.query('COMMIT');
   return send(res,201,{collection:c.rows[0],participantId:p.rows[0].id,collectionToken:token});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){console.error('shared create',e);return send(res,500,{error:'Não foi possível criar a coleta'});}
}
