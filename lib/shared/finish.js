import { db,readJson,send,loggedUser,tokenFrom,hashToken } from './_db.js';
const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const body=await readJson(req),collectionId=String(body.collectionId||'');
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  // Resolve a participação também quando ela já foi finalizada. Isso torna
  // a chamada repetível caso a resposta anterior tenha se perdido na rede.
  const user=await loggedUser(req);
  const token=tokenFrom(req);
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const found=user
    ? await client.query(`SELECT p.id,p.role,p.status,c.status AS collection_status,u.status AS user_status
       FROM cf_participants p JOIN cf_collections c ON c.id=p.collection_id
       LEFT JOIN cf_users u ON u.id=p.user_id
       WHERE p.user_id=$1 AND p.collection_id=$2 FOR UPDATE OF p`,[user.id,collectionId])
    : await client.query(`SELECT p.id,p.role,p.status,c.status AS collection_status,u.status AS user_status
       FROM cf_participants p JOIN cf_collections c ON c.id=p.collection_id
       LEFT JOIN cf_users u ON u.id=p.user_id
       WHERE p.token_hash=$1 AND p.collection_id=$2 FOR UPDATE OF p`,[hashToken(token),collectionId]);
   const p=found.rows[0];
   if(!p){await client.query('ROLLBACK');return send(res,401,{error:'Acesso inválido ou participação não encontrada'});}
   if(p.role==='owner'){await client.query('ROLLBACK');return send(res,403,{error:'O dono acompanha e encerra a coleta; somente usuários finalizam a participação.'});}
   if(p.status==='finished'){
    await client.query('COMMIT');
    return send(res,200,{participant:{id:p.id,status:'finished'},alreadyFinished:true});
   }
   if(p.status!=='active'||p.user_status==='inactive'){
    await client.query('ROLLBACK');
    return send(res,401,{error:'A participação está bloqueada ou o usuário foi desativado'});
   }
   if(p.collection_status!=='open'){
    await client.query('ROLLBACK');
    return send(res,409,{error:'Esta coleta já foi encerrada'});
   }
   const result=await client.query("UPDATE cf_participants SET status='finished',last_seen_at=now() WHERE id=$1 AND status='active' RETURNING id,status,last_seen_at",[p.id]);
   if(!result.rowCount){
    const current=await client.query('SELECT id,status FROM cf_participants WHERE id=$1',[p.id]);
    if(current.rows[0]?.status==='finished'){
     await client.query('COMMIT');
     return send(res,200,{participant:current.rows[0],alreadyFinished:true});
    }
    await client.query('ROLLBACK');
    return send(res,409,{error:'Não foi possível confirmar a finalização. Tente novamente.'});
   }
   await client.query('COMMIT');
   return send(res,200,{participant:result.rows[0],alreadyFinished:false});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){console.error('shared finish',e);return send(res,500,{error:'Não foi possível finalizar sua participação'});}
}
