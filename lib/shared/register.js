import { send } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='POST')return send(res,405,{error:'Método não permitido'});
 return send(res,403,{error:'Cadastro desativado. Solicite acesso ao dono do ColetaFlow.'});
}
