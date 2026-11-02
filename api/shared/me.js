import { send,loggedUser } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{const user=await loggedUser(req);if(!user)return send(res,401,{error:'Sessão expirada'});return send(res,200,{user});}
 catch(e){console.error('shared me',e);return send(res,500,{error:'Não foi possível validar a sessão'});}
}
