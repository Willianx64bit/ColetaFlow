import { db,send,participant } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const collectionId=String(req.query?.collectionId||'');
  if(!/^[0-9a-f-]{36}$/i.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p||p.role!=='user'||p.status!=='active')return send(res,401,{error:'Participação ativa não encontrada'});
  const result=await db().query(`SELECT event_id AS "eventId",barcode,quantity,event_type AS "eventType",device_id AS "deviceId",scanned_at AS "scannedAt",COALESCE(product_description,'') AS description
    FROM cf_participant_draft_events WHERE collection_id=$1 AND participant_id=$2 ORDER BY scanned_at,event_id`,[collectionId,p.id]);
  return send(res,200,{events:result.rows});
 }catch(e){console.error('shared resume',e);return send(res,500,{error:'Não foi possível recuperar as bipagens salvas'});}
}
