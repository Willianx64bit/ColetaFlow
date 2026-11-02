import { db,readJson,send,participant } from './_db.js';

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export default async function handler(req,res) {
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try {
  const body=await readJson(req);
  const collectionId=String(body.collectionId||'');
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p)return send(res,401,{error:'Acesso inválido, bloqueado ou revogado'});
  if(p.collection_status!=='open')return send(res,409,{error:'Esta coleta não está recebendo bipagens'});

  if(!Array.isArray(body.events)||body.events.length===0)return send(res,400,{error:'Nenhuma bipagem enviada'});
  if(body.events.length>500)return send(res,413,{error:'Envie no máximo 500 bipagens por lote'});

  const events=[];
  for(let index=0;index<body.events.length;index++){
   const ev=body.events[index]||{};
   const eventId=String(ev.eventId||'');
   const barcode=String(ev.barcode??'').trim();
   const quantity=Number(ev.quantity);
   const scannedAt=new Date(ev.scannedAt);
   const deviceId=String(ev.deviceId||'');
   const eventType=ev.eventType==='adjustment'?'adjustment':'scan';
   if(!UUID_PATTERN.test(eventId))return send(res,400,{error:`Bipagem ${index+1}: identificador inválido`});
   if(!barcode||barcode.length>120)return send(res,400,{error:`Bipagem ${index+1}: código vazio ou maior que 120 caracteres`});
   if(!Number.isFinite(quantity)||quantity===0||Math.abs(quantity)>1000000||(eventType==='scan'&&quantity<0))return send(res,400,{error:`Bipagem ${index+1}: quantidade inválida`});
   if(Number.isNaN(scannedAt.getTime()))return send(res,400,{error:`Bipagem ${index+1}: data inválida`});
   if(deviceId.length>100)return send(res,400,{error:`Bipagem ${index+1}: identificador do dispositivo muito longo`});
   events.push({eventId,barcode,quantity,eventType,scannedAt,deviceId});
  }

  const client=await db().connect();let inserted=0;
  try {
   await client.query('BEGIN');
   for(const ev of events){
    const result=await client.query(
     'INSERT INTO cf_scan_events(event_id,collection_id,participant_id,barcode,quantity,event_type,device_id,scanned_at) VALUES($1,$2,$3,$4,$5,$6,$7,$8) ON CONFLICT(event_id) DO NOTHING',
     [ev.eventId,collectionId,p.id,ev.barcode,ev.quantity,ev.eventType,ev.deviceId||null,ev.scannedAt]
    );
    if(result.rowCount){
     inserted++;
     continue;
    }
    const existing=await client.query(
     'SELECT collection_id,participant_id,barcode,quantity::text AS quantity,event_type,device_id,scanned_at FROM cf_scan_events WHERE event_id=$1',
     [ev.eventId]
    );
    const saved=existing.rows[0];
    const same=saved
      && saved.collection_id===collectionId
      && saved.participant_id===p.id
      && saved.barcode===ev.barcode
      && Number(saved.quantity)===ev.quantity
      && saved.event_type===ev.eventType
      && (saved.device_id||'')===ev.deviceId
      && new Date(saved.scanned_at).getTime()===ev.scannedAt.getTime();
    if(!same) {
     const conflict=new Error('Identificador de bipagem já utilizado com dados diferentes');
     conflict.statusCode=409;
     throw conflict;
    }
   }
   await client.query('UPDATE cf_participants SET last_seen_at=now() WHERE id=$1',[p.id]);
   await client.query('COMMIT');
  }catch(error){await client.query('ROLLBACK');throw error;}finally{client.release();}
  return send(res,200,{received:events.length,inserted,duplicates:events.length-inserted});
 }catch(error){
  if(error.statusCode===409)return send(res,409,{error:error.message});
  console.error('shared sync',error);
  return send(res,500,{error:'Falha ao sincronizar bipagens'});
 }
}
