import { randomBytes,scryptSync } from 'node:crypto';
import { db,readJson,send,loggedUser } from './_db.js';
export default async function handler(req,res){
 try{
  const owner=await loggedUser(req);if(!owner?.is_owner)return send(res,403,{error:'Apenas o dono pode administrar usuários.'});
  if(req.method==='GET'){
   const r=await db().query(`SELECT u.id,u.email,u.display_name,u.company_name,u.status,u.created_at,
    EXISTS(SELECT 1 FROM cf_platform_owner o WHERE o.user_id=u.id) AS is_owner
    FROM cf_users u WHERE NOT EXISTS(SELECT 1 FROM cf_platform_owner o WHERE o.user_id=u.id) ORDER BY u.created_at DESC`);
   return send(res,200,{users:r.rows});
  }
  if(req.method==='POST'){
   const b=await readJson(req),name=String(b.name||'').trim().slice(0,80),company=String(b.company||'').trim().slice(0,120),email=String(b.email||'').trim().toLowerCase(),password=String(b.password||'');
   if(!name||!company||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||password.length<8||password.length>128)return send(res,400,{error:'Informe nome, empresa, e-mail válido e senha com pelo menos 8 caracteres.'});
   const salt=randomBytes(16).toString('hex'),passwordHash=salt+':'+scryptSync(password,salt,64).toString('hex');
   const r=await db().query("INSERT INTO cf_users(email,password_hash,display_name,company_name,role) VALUES($1,$2,$3,$4,'user') RETURNING id,email,display_name,company_name,status,created_at",[email,passwordHash,name,company]);
   return send(res,201,{user:r.rows[0]});
  }
  if(req.method==='PATCH'){
   const b=await readJson(req),id=String(b.id||''),status=b.status==='disabled'?'disabled':b.status==='active'?'active':'';
   if(!id||!status)return send(res,400,{error:'Usuário ou situação inválida.'});
   const r=await db().query("UPDATE cf_users SET status=$2 WHERE id=$1 AND NOT EXISTS(SELECT 1 FROM cf_platform_owner WHERE user_id=$1) RETURNING id,status",[id,status]);
   if(!r.rowCount)return send(res,404,{error:'Usuário não encontrado.'});
   return send(res,200,{user:r.rows[0]});
  }
  return send(res,405,{error:'Método não permitido'});
 }catch(e){if(e.code==='23505')return send(res,409,{error:'Este e-mail já está cadastrado.'});console.error('shared users',e);return send(res,500,{error:'Não foi possível administrar usuários.'});}
}