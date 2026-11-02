import { db,readJson,send,loggedUser } from './_db.js';

const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export default async function handler(req,res){
 if(req.method!=='PUT')return send(res,405,{error:'Método não permitido'});
 try{
  const owner=await loggedUser(req);
  if(!owner?.is_owner)return send(res,403,{error:'Somente o dono pode definir o mapeamento'});
  const body=await readJson(req),collectionId=String(body.collectionId||''),m=body.mapping||{};
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const sheetName=String(m.sheetName||'').trim();
  const headerRow=Number(m.headerRow);
  const codeColumns=Array.isArray(m.codeColumns)?m.codeColumns.map(Number):[];
  const quantityColumn=Number(m.quantityColumn);
  const descriptionColumn=Number(m.descriptionColumn??-1);
  if(!sheetName||sheetName.length>31||!Number.isInteger(headerRow)||headerRow<0||headerRow>100000)return send(res,400,{error:'Aba ou linha de cabeçalho inválida'});
  if(codeColumns.length<1||codeColumns.length>2||codeColumns.some(x=>!Number.isInteger(x)||x<0||x>10000)||new Set(codeColumns).size!==codeColumns.length)return send(res,400,{error:'Selecione uma ou duas colunas de código válidas'});
  if(!Number.isInteger(quantityColumn)||quantityColumn<0||quantityColumn>10000||codeColumns.includes(quantityColumn))return send(res,400,{error:'Coluna de quantidade inválida'});
  if(!Number.isInteger(descriptionColumn)||descriptionColumn < -1||descriptionColumn>10000)return send(res,400,{error:'Coluna de descrição inválida'});
  const owns=await db().query(`SELECT 1 FROM cf_collections c
    WHERE c.id=$1 AND EXISTS(SELECT 1 FROM cf_participants p WHERE p.collection_id=c.id AND p.user_id=$2 AND p.role='owner' AND p.status='active')`,[collectionId,owner.id]);
  if(!owns.rowCount)return send(res,404,{error:'Coleta não encontrada ou não pertence ao dono'});
  const mapping={sheetName,headerRow,codeColumns,quantityColumn,descriptionColumn};
  const saved=await db().query(`UPDATE cf_collections c SET template_mapping=$2::jsonb
    WHERE c.id=$1 AND (
      NOT EXISTS(SELECT 1 FROM cf_participants p WHERE p.collection_id=c.id AND p.role='user')
      OR c.template_mapping=$2::jsonb
    ) RETURNING c.template_mapping`,[collectionId,JSON.stringify(mapping)]);
  if(!saved.rowCount)return send(res,409,{error:'O mapeamento não pode ser alterado depois que os usuários entram na coleta.'});
  return send(res,200,{ok:true,mapping:saved.rows[0].template_mapping});
 }catch(e){console.error('shared mapping',e);return send(res,500,{error:'Não foi possível salvar o mapeamento'});}
}
