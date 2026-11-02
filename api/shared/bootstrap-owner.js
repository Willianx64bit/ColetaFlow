import { randomBytes,scryptSync,createHash,timingSafeEqual } from 'node:crypto';
import { db,readJson,send } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 try{
  const expected=process.env.COLETAFLOW_OWNER_SETUP_KEY||'',emailAllowed=String(process.env.COLETAFLOW_OWNER_EMAIL||'').trim().toLowerCase();
  const body=await readJson(req),key=String(body.setupKey||''),email=String(body.email||'').trim().toLowerCase(),name=String(body.name||'').trim().slice(0,80),password=String(body.password||'');
  if(!expected||!emailAllowed)return send(res,503,{error:'Configuração inicial do dono ainda não foi habilitada.'});
  const supplied=Buffer.from(key),configured=Buffer.from(expected);
  if(!key||supplied.length!==configured.length||!timingSafeEqual(supplied,configured))return send(res,403,{error:'Chave de configuração inválida'});
  if(email!==emailAllowed||!name||password.length<10||password.length>128)return send(res,400,{error:'Confira o e-mail autorizado, nome e senha (mínimo 10 caracteres).'});
  const client=await db().connect();
  try{
   await client.query('BEGIN');
   const exists=await client.query('SELECT 1 FROM cf_platform_owner LIMIT 1');
   if(exists.rowCount){await client.query('ROLLBACK');return send(res,409,{error:'A conta de dono já foi configurada.'});}
   const salt=randomBytes(16).toString('hex'),passwordHash=salt+':'+scryptSync(password,salt,64).toString('hex');
   const u=await client.query("INSERT INTO cf_users(email,password_hash,display_name,company_name,role) VALUES($1,$2,$3,'ColetaFlow','owner') RETURNING id,email,display_name,company_name,role",[email,passwordHash,name]);
   await client.query('INSERT INTO cf_platform_owner(singleton,user_id) VALUES(true,$1)',[u.rows[0].id]);
   const token=randomBytes(32).toString('base64url'),hash=createHash('sha256').update(token).digest('hex');
   await client.query("INSERT INTO cf_user_sessions(user_id,token_hash,expires_at) VALUES($1,$2,now()+interval '30 days')",[u.rows[0].id,hash]);
   await client.query('COMMIT');return send(res,201,{user:u.rows[0],token});
  }catch(e){await client.query('ROLLBACK');throw e;}finally{client.release();}
 }catch(e){console.error('shared owner bootstrap',e);return send(res,500,{error:'Não foi possível configurar a conta de dono.'});}
}