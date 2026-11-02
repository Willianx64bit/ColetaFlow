import { randomBytes,createHash } from 'node:crypto';
import { db,readJson,send } from './_db.js';
export default async function handler(req,res) {
 if(req.method!=='POST') return send(res,405,{error:'Método não permitido'});
 try {
  const body=await readJson(req), code=String(body.accessCode||'').trim().toUpperCase(), name=String(body.displayName||'').trim().slice(0,60);
  if(!/^[A-F0-9]{8}$/.test(code)||!name) return send(res,400,{error:'Código ou nome inválido'});
  const c=await db().query("SELECT id,title,status FROM cf_collections WHERE access_code=$1",[code]);
  if(!c.rows[0]||c.rows[0].status!=='open') return send(res,404,{error:'Coleta não encontrada ou encerrada'});
  const token=randomBytes(32).toString('base64url'), hash=createHash('sha256').update(token).digest('hex');
  const p=await db().query("INSERT INTO cf_participants(collection_id,display_name,token_hash) VALUES($1,$2,$3) RETURNING id,display_name,role",[c.rows[0].id,name,hash]);
  return send(res,201,{collection:c.rows[0],participant:p.rows[0],token});
 } catch(e){if(e.code==='23505')return send(res,409,{error:'Já existe participante com esse nome nesta coleta'});console.error('shared join',e);return send(res,500,{error:'Não foi possível entrar na coleta'});}
}
