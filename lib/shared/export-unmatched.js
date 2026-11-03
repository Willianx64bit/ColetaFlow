import * as XLSX from 'xlsx';
import { db,send,loggedUser } from './_db.js';

const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const cellText=(cell)=>cell&&cell.v!==undefined&&cell.v!==null?(cell.w!==undefined?String(cell.w).trim():String(cell.v).trim()):'';

export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const owner=await loggedUser(req);
  if(!owner?.is_owner)return send(res,403,{error:'Somente o dono pode exportar os códigos não encontrados'});
  const collectionId=String(req.query?.collectionId||'');
  if(!UUID_PATTERN.test(collectionId))return send(res,400,{error:'Coleta inválida'});
  const result=await db().query(`SELECT c.template_file_name,c.template_bytes,c.template_mapping
    FROM cf_collections c
    WHERE c.id=$1 AND EXISTS(
      SELECT 1 FROM cf_participants p
      WHERE p.collection_id=c.id AND p.user_id=$2 AND p.role='owner' AND p.status='active'
    )`,[collectionId,owner.id]);
  const collection=result.rows[0];
  if(!collection)return send(res,404,{error:'Coleta não encontrada'});
  if(!collection.template_bytes)return send(res,404,{error:'Modelo original não encontrado'});
  if(!collection.template_mapping)return send(res,409,{error:'O dono precisa confirmar o mapeamento da planilha antes de exportar'});
  const mapping=typeof collection.template_mapping==='string'?JSON.parse(collection.template_mapping):collection.template_mapping;
  if(!mapping.sheetName||!Array.isArray(mapping.codeColumns)||!mapping.codeColumns.length)return send(res,409,{error:'Mapeamento da planilha inválido'});
  if(mapping.hasProductList===false)return send(res,409,{error:'Esta coleta foi configurada sem cadastro de produtos; não há comparação de códigos não encontrados.'});
  const original=String(collection.template_file_name||'coleta.xlsx');
  const source=XLSX.read(Buffer.from(collection.template_bytes),{type:'buffer',cellText:true});
  const sheet=source.Sheets[mapping.sheetName];
  if(!sheet)return send(res,409,{error:'A aba definida no mapeamento não existe no modelo original'});
  const range=sheet['!ref']?XLSX.utils.decode_range(sheet['!ref']):{s:{r:0,c:0},e:{r:0,c:0}};
  const headerRow=Number(mapping.headerRow),codeColumns=mapping.codeColumns.map(Number),descriptionColumn=Number(mapping.descriptionColumn);
  if(!Number.isInteger(headerRow)||headerRow<0||codeColumns.some(x=>!Number.isInteger(x)||x<0))return send(res,409,{error:'Mapeamento da planilha inválido'});
  if(codeColumns.some(col=>range.s.c+col>range.e.c))return send(res,409,{error:'As colunas de código não correspondem ao modelo original'});
  const originalCodes=new Set(),descriptions=new Map();
  for(let row=headerRow+1;row<=range.e.r;row++){
   const description=Number.isInteger(descriptionColumn)&&descriptionColumn>=0&&range.s.c+descriptionColumn<=range.e.c?cellText(sheet[XLSX.utils.encode_cell({r:row,c:range.s.c+descriptionColumn})]):'';
   for(const col of codeColumns){const code=cellText(sheet[XLSX.utils.encode_cell({r:row,c:range.s.c+col})]);if(code){originalCodes.add(code);if(!descriptions.has(code))descriptions.set(code,description);}}
  }
  const totals=await db().query(`SELECT t.barcode,t.total_quantity,COALESCE(d.description,'') AS description
    FROM cf_collection_totals t LEFT JOIN (
      SELECT barcode,MAX(product_description) FILTER (WHERE product_description IS NOT NULL AND product_description<>'') AS description
      FROM cf_scan_events WHERE collection_id=$1 GROUP BY barcode
    ) d ON d.barcode=t.barcode
    WHERE t.collection_id=$1 ORDER BY t.barcode`,[collectionId]);
  const missing=totals.rows.filter(item=>!originalCodes.has(String(item.barcode)));
  if(!missing.length)return send(res,409,{error:'Nenhum código não encontrado foi registrado nesta coleta.'});
  const rows=[['Código lido','Descrição','Quantidade coletada'],...missing.map(item=>[String(item.barcode),String(item.description||''),Number(item.total_quantity)])];
  const report=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(report,XLSX.utils.aoa_to_sheet(rows),'Não encontrados');
  const baseName=original.replace(/\.(xlsx|xls|csv)$/i,'');
  return send(res,200,{fileName:`${baseName}-nao-encontrados.xlsx`,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',base64:XLSX.write(report,{bookType:'xlsx',type:'base64'}),unmatchedCodes:missing.length,totalUnits:missing.reduce((sum,item)=>sum+Number(item.total_quantity||0),0)});
 }catch(e){console.error('shared export unmatched',e);return send(res,500,{error:'Não foi possível gerar a planilha de códigos não encontrados'});}
}
