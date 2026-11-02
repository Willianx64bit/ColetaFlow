import { db,readJson,send,loggedUser } from './_db.js';
const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const owner=await loggedUser(req);
  if(!owner?.is_owner)return send(res,403,{error:'Somente o dono pode encerrar a coleta'});
  const body=await readJson(req),collectionId=String(body.collectionId||'');
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const collection=await client.query('SELECT id,status FROM cf_collections WHERE id=$1 FOR UPDATE',[collectionId]);
   if(!collection.rowCount){await client.query('ROLLBACK');return send(res,404,{error:'Coleta não encontrada'});}
   if(collection.rows[0].status!=='open'){await client.query('ROLLBACK');return send(res,409,{error:'Esta coleta já foi encerrada ou não está aberta'});}
   const ownerParticipant=await client.query("SELECT 1 FROM cf_participants WHERE collection_id=$1 AND user_id=$2 AND role='owner' AND status='active'",[collectionId,owner.id]);
   if(!ownerParticipant.rowCount){await client.query('ROLLBACK');return send(res,404,{error:'Coleta não encontrada'});}
   const active=await client.query("SELECT COUNT(*)::int AS count FROM cf_participants WHERE collection_id=$1 AND role='user' AND status='active'",[collectionId]);
   if(Number(active.rows[0]?.count||0)>0){await client.query('ROLLBACK');return send(res,409,{error:`Ainda existem ${active.rows[0].count} usuário(s) trabalhando. Aguarde a finalização de todos.`,activeUsers:Number(active.rows[0].count)});}
   const result=await client.query("UPDATE cf_collections SET status='completed',completed_at=now() WHERE id=$1 RETURNING id,status,completed_at",[collectionId]);
   await client.query('COMMIT');
   return send(res,200,{collection:result.rows[0]});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){console.error('shared close',e);return send(res,500,{error:'Não foi possível encerrar a coleta'});}
}
