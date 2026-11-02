import { randomBytes,createHash } from 'node:crypto';
import { db,readJson,send,loggedUser } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const user=await loggedUser(req);if(!user)return send(res,401,{error:'Entre na sua conta para participar'});
  const body=await readJson(req),code=String(body.accessCode||'').trim().toUpperCase();
  if(!/^[A-F0-9]{8}$/.test(code))return send(res,400,{error:'Código de coleta inválido'});
  const c=await db().query("SELECT id,title,status FROM cf_collections WHERE access_code=$1",[code]);
  if(!c.rows[0]||c.rows[0].status!=='open')return send(res,404,{error:'Coleta não encontrada ou encerrada'});
  const token=randomBytes(32).toString('base64url'),hash=createHash('sha256').update(token).digest('hex');
  const p=await db().query("INSERT INTO cf_participants(collection_id,display_name,token_hash,role,user_id) VALUES($1,$2,$3,'collector',$4) RETURNING id,display_name,role",[c.rows[0].id,user.display_name,hash,user.id]);
  return send(res,201,{collection:c.rows[0],participant:p.rows[0],collectionToken:token});
 }catch(e){if(e.code==='23505')return send(res,409,{error:'Você já participa desta coleta ou seu nome já está em uso'});console.error('shared join',e);return send(res,500,{error:'Não foi possível entrar na coleta'});}
}
