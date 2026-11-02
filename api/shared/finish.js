import { db,readJson,send,participant } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const body=await readJson(req),collectionId=String(body.collectionId||'');
  if(!/^[0-9a-f-]{36}$/i.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p)return send(res,401,{error:'Acesso inválido ou revogado'});
  if(p.collection_status!=='open')return send(res,409,{error:'Esta coleta já foi encerrada'});
  const result=await db().query("UPDATE cf_participants SET status='finished',last_seen_at=now() WHERE id=$1 AND status='active' RETURNING id,status,last_seen_at",[p.id]);
  if(!result.rowCount)return send(res,409,{error:'Esta participação já foi finalizada'});
  return send(res,200,{participant:result.rows[0]});
 }catch(e){console.error('shared finish',e);return send(res,500,{error:'Não foi possível finalizar sua participação'});}
}
