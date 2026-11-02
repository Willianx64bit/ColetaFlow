import { db,send,participant } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const collectionId=String(req.query?.collectionId||'');
  if(!/^[0-9a-f-]{36}$/i.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p)return send(res,401,{error:'Acesso inválido ou revogado'});
  const [collection,totals,people]=await Promise.all([
   db().query('SELECT id,title,access_code,status,created_at FROM cf_collections WHERE id=$1',[collectionId]),
   db().query('SELECT barcode,total_quantity,scan_count,last_received_at FROM cf_collection_totals WHERE collection_id=$1 ORDER BY barcode',[collectionId]),
   db().query("SELECT id,display_name,role,status,joined_at,last_seen_at FROM cf_participants WHERE collection_id=$1 ORDER BY joined_at",[collectionId])
  ]);
  if(!collection.rows[0])return send(res,404,{error:'Coleta não encontrada'});
  return send(res,200,{collection:collection.rows[0],totals:totals.rows,participants:people.rows});
 }catch(e){console.error('shared progress',e);return send(res,500,{error:'Não foi possível carregar o progresso'});}
}
