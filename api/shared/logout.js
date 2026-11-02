import { db,send,tokenFrom,hashToken } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const token=tokenFrom(req);
  if(!token)return send(res,401,{error:'Sessão não informada'});
  await db().query('DELETE FROM cf_user_sessions WHERE token_hash=$1',[hashToken(token)]);
  return send(res,200,{ok:true});
 }catch(e){console.error('shared logout',e);return send(res,500,{error:'Não foi possível encerrar a sessão'});}
}
