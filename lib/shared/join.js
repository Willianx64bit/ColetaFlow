import { randomBytes,createHash } from 'node:crypto';
import { db,readJson,send,loggedUser } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const user=await loggedUser(req);if(!user)return send(res,401,{error:'Entre na sua conta para participar'});
  const body=await readJson(req),code=String(body.accessCode||'').trim().toUpperCase();
  if(!/^[A-F0-9]{8}$/.test(code))return send(res,400,{error:'Código de coleta inválido'});
  const token=randomBytes(32).toString('base64url'),hash=createHash('sha256').update(token).digest('hex');
  const role=user.is_owner?'owner':'user';
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const c=await client.query('SELECT id,title,status,template_mapping FROM cf_collections WHERE access_code=$1 FOR UPDATE',[code]);
   if(!c.rows[0]||c.rows[0].status!=='open'){await client.query('ROLLBACK');return send(res,404,{error:'Coleta não encontrada ou encerrada'});}
   if(!c.rows[0].template_mapping){await client.query('ROLLBACK');return send(res,409,{error:'O dono ainda precisa confirmar o mapeamento da planilha antes de liberar a coleta'});}
   const p=await client.query(`INSERT INTO cf_participants(collection_id,display_name,token_hash,role,user_id)
     VALUES($1,$2,$3,$4,$5)
     ON CONFLICT(collection_id,user_id) WHERE user_id IS NOT NULL
     DO UPDATE SET display_name=EXCLUDED.display_name,token_hash=EXCLUDED.token_hash,
       role=EXCLUDED.role,status='active',last_seen_at=now()
     RETURNING id,display_name,role,status`,[c.rows[0].id,user.display_name,hash,role,user.id]);
   if(!p.rowCount){await client.query('ROLLBACK');return send(res,409,{error:'Não foi possível reativar sua participação nesta coleta'});}
   await client.query('COMMIT');
   return send(res,200,{collection:c.rows[0],participant:p.rows[0],collectionToken:token});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){if(e.code==='23505')return send(res,409,{error:'Não foi possível registrar sua participação'});console.error('shared join',e);return send(res,500,{error:'Não foi possível entrar na coleta'});}
}
