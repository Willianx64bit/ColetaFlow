import { db,send,participant } from './_db.js';
export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const collectionId=String(req.query?.collectionId||'');
  if(!/^[0-9a-f-]{36}$/i.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const p=await participant(req,collectionId);
  if(!p)return send(res,401,{error:'Acesso inválido, bloqueado ou revogado'});
  const result=await db().query('SELECT template_file_name,template_mime_type,template_bytes,template_sha256,template_mapping FROM cf_collections WHERE id=$1',[collectionId]);
  const row=result.rows[0];
  if(!row||!row.template_bytes)return send(res,404,{error:'Modelo original não encontrado'});
  return send(res,200,{
   fileName:row.template_file_name,
   mimeType:row.template_mime_type,
   sha256:row.template_sha256,
   mapping:row.template_mapping,
   base64:Buffer.from(row.template_bytes).toString('base64')
  });
 }catch(e){console.error('shared template',e);return send(res,500,{error:'Não foi possível carregar o modelo original'});}
}
