import { randomBytes,createHash } from 'node:crypto';
import { db,readJson,send,loggedUser } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const user=await loggedUser(req);if(!user)return send(res,401,{error:'Entre na sua conta para criar uma coleta'});
  if(!user.is_owner)return send(res,403,{error:'Somente o dono pode criar coletas'});
  const body=await readJson(req),title=String(body.title||'Coleta compartilhada').trim().slice(0,120);
  const templateFileName=String(body.templateFileName||'').trim().slice(0,255);
  const templateMimeType=String(body.templateMimeType||'application/octet-stream').slice(0,120);
  const templateBase64=String(body.templateBase64||'');
  if(!title)return send(res,400,{error:'Informe o nome da coleta'});
  if(!/\\.(xlsx|xls|csv)$/i.test(templateFileName)||!templateBase64)return send(res,400,{error:'Envie o modelo XLSX, XLS ou CSV da coleta'});
  if(!/^[A-Za-z0-9+/]+={0,2}$/.test(templateBase64)||templateBase64.length%4!==0)return send(res,400,{error:'Arquivo do modelo inválido'});
  if(templateBase64.length>4*1024*1024)return send(res,413,{error:'O modelo deve ter no máximo 3 MB'});
  const templateBytes=Buffer.from(templateBase64,'base64');
  if(!templateBytes.length||templateBytes.length>3*1024*1024)return send(res,413,{error:'O modelo está vazio ou ultrapassa 3 MB'});
  const templateSha256=createHash('sha256').update(templateBytes).digest('hex');
  const code=randomBytes(4).toString('hex').toUpperCase(),token=randomBytes(32).toString('base64url'),hash=createHash('sha256').update(token).digest('hex');
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const c=await client.query("INSERT INTO cf_collections(access_code,title,template_file_name,template_mime_type,template_bytes,template_sha256) VALUES($1,$2,$3,$4,$5,$6) RETURNING id,access_code,title,status,created_at,template_file_name,template_mime_type,template_sha256",[code,title,templateFileName,templateMimeType,templateBytes,templateSha256]);
   const p=await client.query("INSERT INTO cf_participants(collection_id,display_name,token_hash,role,user_id) VALUES($1,$2,$3,'owner',$4) RETURNING id",[c.rows[0].id,user.display_name,hash,user.id]);
   await client.query('COMMIT');
   return send(res,201,{collection:c.rows[0],participantId:p.rows[0].id,collectionToken:token});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){console.error('shared create',e);return send(res,500,{error:'Não foi possível criar a coleta'});}
}
