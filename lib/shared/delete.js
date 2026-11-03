import { db,readJson,send,loggedUser } from './_db.js';
const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const owner=await loggedUser(req);
  if(!owner?.is_owner)return send(res,403,{error:'Somente o dono pode apagar coletas.'});
  const body=await readJson(req),collectionId=String(body.collectionId||'');
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida.'});
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const owned=await client.query(`SELECT c.id FROM cf_collections c
     JOIN cf_participants p ON p.collection_id=c.id
     WHERE c.id=$1 AND p.user_id=$2 AND p.role='owner' AND p.status='active'
     FOR UPDATE OF c`,[collectionId,owner.id]);
   if(!owned.rowCount){await client.query('ROLLBACK');return send(res,404,{error:'Coleta não encontrada.'});}
   await client.query('DELETE FROM cf_scan_events WHERE collection_id=$1',[collectionId]);
   await client.query('DELETE FROM cf_participants WHERE collection_id=$1',[collectionId]);
   await client.query('DELETE FROM cf_collections WHERE id=$1',[collectionId]);
   await client.query('COMMIT');
   return send(res,200,{deleted:true,collectionId});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){console.error('shared delete collection',e);return send(res,500,{error:'Não foi possível apagar a coleta.'});}
}
