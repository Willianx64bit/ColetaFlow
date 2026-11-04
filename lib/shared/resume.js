import { db,send,participant } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const collectionId=String(req.query?.collectionId||'');
  if(!/^[0-9a-f-]{36}$/i.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p||p.role!=='user'||p.status!=='active')return send(res,401,{error:'Participação ativa não encontrada'});
  // Recover both private drafts and events already accepted by the legacy
  // finish/sync flow. This keeps older sessions resumable after an upgrade.
  const result=await db().query(`SELECT event_id AS "eventId",barcode,quantity,event_type AS "eventType",device_id AS "deviceId",scanned_at AS "scannedAt",COALESCE(product_description,'') AS description
    FROM cf_participant_draft_events WHERE collection_id=$1 AND participant_id=$2
    UNION
    SELECT e.event_id AS "eventId",e.barcode,e.quantity,e.event_type AS "eventType",e.device_id AS "deviceId",e.scanned_at AS "scannedAt",COALESCE(e.product_description,'') AS description
    FROM cf_scan_events e
    WHERE e.collection_id=$1 AND e.participant_id=$2
      AND NOT EXISTS (SELECT 1 FROM cf_participant_draft_events d WHERE d.event_id=e.event_id)
    ORDER BY "scannedAt","eventId"`,[collectionId,p.id]);
  return send(res,200,{events:result.rows});
 }catch(e){console.error('shared resume',e);return send(res,500,{error:'Não foi possível recuperar as bipagens salvas'});}
}
