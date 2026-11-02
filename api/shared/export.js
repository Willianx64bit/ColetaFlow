import * as XLSX from 'xlsx';
import { db,send,loggedUser } from './_db.js';

const UUID_PATTERN=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const cellText=(cell)=>cell&&cell.v!==undefined&&cell.v!==null?(cell.w!==undefined?String(cell.w).trim():String(cell.v).trim()):'';

export default async function handler(req,res){
 if(req.method!=='GET')return send(res,405,{error:'Método não permitido'});
 try{
  const owner=await loggedUser(req);
  if(!owner?.is_owner)return send(res,403,{error:'Somente o dono pode exportar a planilha consolidada'});
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
  const workbook=XLSX.read(Buffer.from(collection.template_bytes),{type:'buffer',cellStyles:true,cellNF:true});
  const sheet=workbook.Sheets[mapping.sheetName];
  if(!sheet)return send(res,409,{error:'A aba definida no mapeamento não existe no modelo original'});
  const range=sheet['!ref']?XLSX.utils.decode_range(sheet['!ref']):{s:{r:0,c:0},e:{r:0,c:0}};
  const headerRow=Number(mapping.headerRow),codeColumns=mapping.codeColumns.map(Number),quantityColumn=Number(mapping.quantityColumn);
  if(!Number.isInteger(headerRow)||headerRow<0||!codeColumns.length||!Number.isInteger(quantityColumn)||quantityColumn<0)return send(res,409,{error:'Mapeamento da planilha inválido'});
  const quantityCol=range.s.c+quantityColumn;
  if(quantityCol>range.e.c||codeColumns.some(col=>range.s.c+col>range.e.c))return send(res,409,{error:'As colunas salvas não correspondem mais ao modelo original'});
  for(let row=headerRow+1;row<=range.e.r;row++){
    const address=XLSX.utils.encode_cell({r:row,c:quantityCol});
    const previous=sheet[address];
    sheet[address]={t:'s',v:'',...(previous?.s?{s:previous.s}:{})};
  }
  const totals=await db().query('SELECT barcode,total_quantity FROM cf_collection_totals WHERE collection_id=$1 ORDER BY barcode',[collectionId]);
  let matched=0,unmatched=0,units=0;
  for(const item of totals.rows){
    const barcode=String(item.barcode),quantity=Number(item.total_quantity);
    let found=false;
    for(let row=headerRow+1;row<=range.e.r;row++){
      const codes=codeColumns.map(col=>cellText(sheet[XLSX.utils.encode_cell({r:row,c:range.s.c+col})]));
      if(codes.includes(barcode)){
        const address=XLSX.utils.encode_cell({r:row,c:quantityCol});
        const previous=sheet[address];
        sheet[address]={t:'n',v:quantity,...(previous?.s?{s:previous.s}:{})};
        found=true;break;
      }
    }
    if(found){matched++;units+=quantity;}else unmatched++;
  }
  const original=String(collection.template_file_name||'coleta.xlsx');
  const extension=(original.match(/\.(xlsx|xls|csv)$/i)||[])[1]?.toLowerCase()||'xlsx';
  const bookType=extension==='csv'?'csv':extension==='xls'?'biff8':'xlsx';
  const output=XLSX.write(workbook,{bookType,type:'base64',sheet:mapping.sheetName,cellStyles:true});
  const baseName=original.replace(/\.(xlsx|xls|csv)$/i,'');
  const outputExtension=extension==='xls'?'xls':extension;
  const mimeType=extension==='csv'?'text/csv':extension==='xls'?'application/vnd.ms-excel':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  return send(res,200,{fileName:`${baseName}-coletado.${outputExtension}`,mimeType,base64:output,matchedCodes:matched,unmatchedCodes:unmatched,totalUnits:units});
 }catch(e){console.error('shared export',e);return send(res,500,{error:'Não foi possível gerar a planilha consolidada'});}
}
