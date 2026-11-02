import { randomBytes,scryptSync,createHash } from 'node:crypto';
import { db,readJson,send } from './_db.js';
const emailOk=v=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const b=await readJson(req),name=String(b.name||'').trim().slice(0,80),company=String(b.company||'').trim().slice(0,120),email=String(b.email||'').trim().toLowerCase(),password=String(b.password||'');
  if(!name||!company||!emailOk(email)||password.length<8||password.length>128)return send(res,400,{error:'Preencha nome, empresa, e-mail válido e senha com pelo menos 8 caracteres'});
  const salt=randomBytes(16).toString('hex'),passwordHash=salt+':'+scryptSync(password,salt,64).toString('hex');
  const token=randomBytes(32).toString('base64url'),tokenHash=createHash('sha256').update(token).digest('hex');
  const c=await db().query('INSERT INTO cf_users(email,password_hash,display_name,company_name,role) VALUES($1,$2,$3,$4,$5) RETURNING id,email,display_name,company_name,role',[email,passwordHash,name,company,'manager']);
  await db().query("INSERT INTO cf_user_sessions(user_id,token_hash,expires_at) VALUES($1,$2,now()+interval '30 days')",[c.rows[0].id,tokenHash]);
  return send(res,201,{user:c.rows[0],token});
 }catch(e){if(e.code==='23505')return send(res,409,{error:'Este e-mail já possui cadastro'});console.error('shared register',e);return send(res,500,{error:'Não foi possível criar a conta'});}
}
