import { db,send,loggedUser } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const owner=await loggedUser(req);
  if(!owner?.is_owner)return send(res,403,{error:'Somente o dono pode consultar suas coletas.'});
  const result=await db().query(`SELECT c.id,c.title,c.access_code,c.status,c.created_at,c.completed_at,
    COALESCE(p.participant_count,0)::int AS participant_count,
    COALESCE(p.active_count,0)::int AS active_count,
    COALESCE(p.finished_count,0)::int AS finished_count,
    COALESCE(e.event_count,0)::int AS event_count,
    COALESCE(e.total_units,0)::numeric AS total_units,
    e.last_scan_at
   FROM cf_collections c
   JOIN cf_participants owner_participant
     ON owner_participant.collection_id=c.id
    AND owner_participant.user_id=$1
    AND owner_participant.role='owner'
    AND owner_participant.status='active'
   LEFT JOIN (
     SELECT collection_id,
       COUNT(*) FILTER (WHERE role='user' AND (user_id IS NULL OR user_id NOT IN (SELECT user_id FROM cf_platform_owner))) AS participant_count,
       COUNT(*) FILTER (WHERE role='user' AND status='active' AND (user_id IS NULL OR user_id NOT IN (SELECT user_id FROM cf_platform_owner))) AS active_count,
       COUNT(*) FILTER (WHERE role='user' AND status='finished' AND (user_id IS NULL OR user_id NOT IN (SELECT user_id FROM cf_platform_owner))) AS finished_count
     FROM cf_participants GROUP BY collection_id
   ) p ON p.collection_id=c.id
   LEFT JOIN (
     SELECT collection_id,COUNT(*) AS event_count,SUM(quantity) AS total_units,MAX(scanned_at) AS last_scan_at
     FROM cf_scan_events GROUP BY collection_id
   ) e ON e.collection_id=c.id
   ORDER BY c.created_at DESC`,[owner.id]);
  return send(res,200,{collections:result.rows});
 }catch(e){console.error('shared collections list',e);return send(res,500,{error:'Não foi possível carregar suas coletas.'});}
}
