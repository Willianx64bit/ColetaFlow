import { db,readJson,send,participant } from './_db.js';
export default async function handler(req,res) {
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try {
  const body=await readJson(req);
  const collectionId=String(body.collectionId||'');
  const p=await participant(req,collectionId);
  if(!p)return send(res,401,{error:'Acesso inválido ou revogado'});
  if(p.collection_status!=='open')return send(res,409,{error:'Esta coleta não está recebendo bipagens'});
  const events=Array.isArray(body.events)?body.events.slice(0,500):[];
  if(!events.length)return send(res,400,{error:'Nenhuma bipagem enviada'});
  const client=await db().connect();let inserted=0;
  try {
   await client.query('BEGIN');
   for(const ev of events){
    const eventId=String(ev.eventId||'');
    const barcode=String(ev.barcode||'').trim().slice(0,120);
    const quantity=Number(ev.quantity);
    const scannedAt=new Date(ev.scannedAt||Date.now());
    if(!/^[0-9a-f-]{36}$/i.test(eventId)||!barcode||!Number.isFinite(quantity)||quantity<=0||quantity>1000000||Number.isNaN(scannedAt.getTime()))continue;
    const result=await client.query(
     'INSERT INTO cf_scan_events(event_id,collection_id,participant_id,barcode,quantity,device_id,scanned_at) VALUES($1,$2,$3,$4,$5,$6,$7) ON CONFLICT(event_id) DO NOTHING',
     [eventId,collectionId,p.id,barcode,quantity,String(ev.deviceId||'').slice(0,100),scannedAt]
    );
    inserted+=result.rowCount||0;
   }
   await client.query('UPDATE cf_participants SET last_seen_at=now() WHERE id=$1',[p.id]);
   await client.query('COMMIT');
  }catch(error){await client.query('ROLLBACK');throw error;}finally{client.release();}
  return send(res,200,{received:events.length,inserted,duplicates:events.length-inserted});
 }catch(error){console.error('shared sync',error);return send(res,500,{error:'Falha ao sincronizar bipagens'});}
}