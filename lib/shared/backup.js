import { db,readJson,send,participant } from './_db.js';
const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const body=await readJson(req),collectionId=String(body.collectionId||'');
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p||p.role!=='user'||p.status!=='active')return send(res,401,{error:'Participação de usuário ativa não encontrada'});
  if(p.collection_status!=='open')return send(res,409,{error:'Esta coleta não está recebendo cópias de segurança'});
  if(!Array.isArray(body.events)||body.events.length===0)return send(res,400,{error:'Nenhuma bipagem enviada'});
  if(body.events.length>500)return send(res,413,{error:'Envie no máximo 500 bipagens por lote'});
  const client=await db().connect();let saved=0;
  try{
   await client.query('BEGIN');
   const active=await client.query(`SELECT p.id FROM cf_participants p JOIN cf_collections c ON c.id=p.collection_id
     LEFT JOIN cf_users u ON u.id=p.user_id
     WHERE p.id=$1 AND p.collection_id=$2 AND p.status='active' AND c.status='open' AND (p.user_id IS NULL OR u.status='active')
     FOR UPDATE OF p`,[p.id,collectionId]);
   if(!active.rowCount){const e=new Error('Participação bloqueada ou coleta encerrada');e.statusCode=409;throw e;}
   for(let i=0;i<body.events.length;i++){
    const ev=body.events[i]||{},eventId=String(ev.eventId||''),barcode=String(ev.barcode??'').trim(),quantity=Number(ev.quantity);
    const eventType=ev.eventType===undefined?'scan':ev.eventType,scannedAt=new Date(ev.scannedAt),deviceId=String(ev.deviceId||''),description=String(ev.description||'').trim();
    if(!UUID_PATTERN.test(eventId)||!barcode||barcode.length>120||(eventType!=='scan'&&eventType!=='adjustment')||!Number.isFinite(quantity)||quantity===0||Math.abs(quantity)>1000000||(eventType==='scan'&&quantity<0)||Number.isNaN(scannedAt.getTime())||deviceId.length>100||description.length>180)
      return await rollbackBad(client,res,400,{error:`Bipagem ${i+1}: dados inválidos`});
    const r=await client.query(`INSERT INTO cf_participant_draft_events(event_id,collection_id,participant_id,barcode,quantity,event_type,device_id,scanned_at,product_description)
      VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9) ON CONFLICT(event_id) DO NOTHING`,
      [eventId,collectionId,p.id,barcode,quantity,eventType,deviceId||null,scannedAt,description||null]);
    saved+=r.rowCount;
   }
   await client.query('COMMIT');
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
  return send(res,200,{received:body.events.length,saved});
 }catch(e){if(e.statusCode===409)return send(res,409,{error:e.message});console.error('shared backup',e);return send(res,500,{error:'Não foi possível salvar a cópia de segurança das bipagens'});}
}
async function rollbackBad(client,res,status,body){await client.query('ROLLBACK');return send(res,status,body);}
