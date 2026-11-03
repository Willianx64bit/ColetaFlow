import { db,send,participant } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const collectionId=String(req.query?.collectionId||'');
  if(!/^[0-9a-f-]{36}$/i.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p)return send(res,401,{error:'Acesso inválido ou revogado'});
  const [collection,totals,people]=await Promise.all([
   db().query('SELECT id,title,access_code,status,created_at,completed_at FROM cf_collections WHERE id=$1',[collectionId]),
   db().query('SELECT barcode,total_quantity,scan_count,last_received_at FROM cf_collection_totals WHERE collection_id=$1 ORDER BY barcode',[collectionId]),
   db().query(`SELECT p.id,p.display_name,CASE WHEN EXISTS(SELECT 1 FROM cf_platform_owner o WHERE o.user_id=p.user_id) THEN 'owner' ELSE p.role END AS role,p.status,p.joined_at,p.last_seen_at,
      COALESCE(s.event_count,0)::int AS event_count,
      COALESCE(s.total_units,0)::numeric AS total_units,
      COALESCE(s.distinct_codes,0)::int AS distinct_codes,
      s.last_scan_at
    FROM cf_participants p
    LEFT JOIN (
      SELECT participant_id,COUNT(*) AS event_count,SUM(quantity) AS total_units,
        COUNT(DISTINCT barcode) AS distinct_codes,MAX(scanned_at) AS last_scan_at
      FROM cf_scan_events WHERE collection_id=$1 GROUP BY participant_id
    ) s ON s.participant_id=p.id
    WHERE p.collection_id=$1 AND ($2::boolean OR p.id=$3)
    ORDER BY p.joined_at`,[collectionId,p.role==='owner',p.id])
  ]);
  if(!collection.rows[0])return send(res,404,{error:'Coleta não encontrada'});
  return send(res,200,{collection:collection.rows[0],totals:totals.rows,participants:people.rows});
 }catch(e){console.error('shared progress',e);return send(res,500,{error:'Não foi possível carregar o progresso'});}
}
