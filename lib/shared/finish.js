import { db,readJson,send,participant } from './_db.js';
const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const body=await readJson(req),collectionId=String(body.collectionId||'');
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p)return send(res,401,{error:'Acesso inválido ou revogado'});
  if(p.role==='owner')return send(res,403,{error:'O dono acompanha e encerra a coleta; somente usuários finalizam a participação.'});
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const collection=await client.query('SELECT status FROM cf_collections WHERE id=$1 FOR UPDATE',[collectionId]);
   if(!collection.rowCount||collection.rows[0].status!=='open'){await client.query('ROLLBACK');return send(res,409,{error:'Esta coleta já foi encerrada'});}
   const active=await client.query(`SELECT p.id FROM cf_participants p LEFT JOIN cf_users u ON u.id=p.user_id
     WHERE p.id=$1 AND p.collection_id=$2 AND p.status='active' AND (p.user_id IS NULL OR u.status='active')
     FOR UPDATE OF p`,[p.id,collectionId]);
   if(!active.rowCount){await client.query('ROLLBACK');return send(res,401,{error:'A participação foi bloqueada ou já finalizada'});}
   const result=await client.query("UPDATE cf_participants SET status='finished',last_seen_at=now() WHERE id=$1 RETURNING id,status,last_seen_at",[p.id]);
   await client.query('COMMIT');
   return send(res,200,{participant:result.rows[0]});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){console.error('shared finish',e);return send(res,500,{error:'Não foi possível finalizar sua participação'});}
}
