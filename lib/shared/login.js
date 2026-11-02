import { randomBytes,scryptSync,timingSafeEqual,createHash } from 'node:crypto';
import { db,readJson,send } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const b=await readJson(req),email=String(b.email||'').trim().toLowerCase(),password=String(b.password||'');
  const r=await db().query('SELECT u.id,u.email,u.password_hash,u.display_name,u.company_name,u.role,u.status,EXISTS(SELECT 1 FROM cf_platform_owner o WHERE o.user_id=u.id) AS is_owner FROM cf_users u WHERE u.email=$1',[email]);
  const u=r.rows[0]; if(!u||u.status!=='active')return send(res,401,{error:'E-mail ou senha inválidos'});
  const [salt,stored]=u.password_hash.split(':');
  const candidate=scryptSync(password,salt,64),expected=Buffer.from(stored,'hex');
  if(candidate.length!==expected.length||!timingSafeEqual(candidate,expected))return send(res,401,{error:'E-mail ou senha inválidos'});
  const token=randomBytes(32).toString('base64url'),hash=createHash('sha256').update(token).digest('hex');
  await db().query("INSERT INTO cf_user_sessions(user_id,token_hash,expires_at) VALUES($1,$2,now()+interval '30 days')",[u.id,hash]);
  delete u.password_hash;
  return send(res,200,{user:{id:u.id,email:u.email,name:u.display_name,company:u.company_name,role:u.role,is_owner:u.is_owner},token});
 }catch(e){console.error('shared login',e);return send(res,500,{error:'Não foi possível entrar'});}
}
