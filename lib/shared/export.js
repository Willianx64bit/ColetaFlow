import ExcelJS from 'exceljs';
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
      WHERE p.collection_id=c.id AND p.user_id=$2 AND p.role='owner' AND p.status IN ('active','finished')
    )`,[collectionId,owner.id]);
  const collection=result.rows[0];
  if(!collection)return send(res,404,{error:'Coleta não encontrada'});
  if(!collection.template_bytes)return send(res,404,{error:'Modelo original não encontrado'});
  if(!collection.template_mapping)return send(res,409,{error:'O dono precisa confirmar o mapeamento da planilha antes de exportar'});
  const mapping=typeof collection.template_mapping==='string'?JSON.parse(collection.template_mapping):collection.template_mapping;
  if(!mapping.sheetName||!Array.isArray(mapping.codeColumns)||!mapping.codeColumns.length)return send(res,409,{error:'Mapeamento da planilha inválido'});
  const totals=await db().query('SELECT barcode,total_quantity FROM cf_collection_totals WHERE collection_id=$1 ORDER BY barcode',[collectionId]);
  const original=String(collection.template_file_name||'coleta.xlsx');
  const extension=(original.match(/\.(xlsx|xls|csv)$/i)||[])[1]?.toLowerCase()||'xlsx';
  let output,mimeType;
  let matched=0,unmatched=0,units=0;

  if(extension==='xlsx'){
    const workbook=new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(collection.template_bytes));
    const worksheet=workbook.getWorksheet(mapping.sheetName);
    if(!worksheet)return send(res,409,{error:'A aba definida no mapeamento não existe no modelo original'});
    const leftColumn=worksheet.dimensions?.left||1;
    const headerRow=Number(mapping.headerRow),quantityColumn=Number(mapping.quantityColumn);
    const codeColumns=mapping.codeColumns.map(Number);
    if(!Number.isInteger(headerRow)||headerRow<0||!Number.isInteger(quantityColumn)||quantityColumn<0||codeColumns.some(x=>!Number.isInteger(x)||x<0))return send(res,409,{error:'Mapeamento da planilha inválido'});
    const quantityAbsolute=leftColumn+quantityColumn;
    if(quantityAbsolute>worksheet.columnCount||codeColumns.some(col=>leftColumn+col>worksheet.columnCount))return send(res,409,{error:'As colunas salvas não correspondem ao modelo original'});
    const firstDataRow=headerRow+2;
    for(let row=firstDataRow;row<=worksheet.rowCount;row++)worksheet.getCell(row,quantityAbsolute).value=null;
    for(const item of totals.rows){
      const barcode=String(item.barcode),quantity=Number(item.total_quantity);
      let found=false;
      for(let row=firstDataRow;row<=worksheet.rowCount;row++){
        const codes=codeColumns.map(col=>String(worksheet.getCell(row,leftColumn+col).text||'').trim());
        if(codes.includes(barcode)){
          worksheet.getCell(row,quantityAbsolute).value=quantity;
          found=true;break;
        }
      }
      if(found){matched++;units+=quantity;}else unmatched++;
    }
    output=Buffer.from(await workbook.xlsx.writeBuffer()).toString('base64');
    mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  }else{
    const workbook=XLSX.read(Buffer.from(collection.template_bytes),{type:'buffer',cellStyles:true,cellNF:true});
    const sheet=workbook.Sheets[mapping.sheetName];
    if(!sheet)return send(res,409,{error:'A aba definida no mapeamento não existe no modelo original'});
    const range=sheet['!ref']?XLSX.utils.decode_range(sheet['!ref']):{s:{r:0,c:0},e:{r:0,c:0}};
    const headerRow=Number(mapping.headerRow),codeColumns=mapping.codeColumns.map(Number),quantityColumn=Number(mapping.quantityColumn);
    if(!Number.isInteger(headerRow)||headerRow<0||!Number.isInteger(quantityColumn)||quantityColumn<0)return send(res,409,{error:'Mapeamento da planilha inválido'});
    const quantityCol=range.s.c+quantityColumn;
    if(quantityCol>range.e.c||codeColumns.some(col=>range.s.c+col>range.e.c))return send(res,409,{error:'As colunas salvas não correspondem mais ao modelo original'});
    for(let row=headerRow+1;row<=range.e.r;row++){
      const address=XLSX.utils.encode_cell({r:row,c:quantityCol}),previous=sheet[address];
      sheet[address]={t:'s',v:'',...(previous?.s?{s:previous.s}:{})};
    }
    for(const item of totals.rows){
      const barcode=String(item.barcode),quantity=Number(item.total_quantity);
      let found=false;
      for(let row=headerRow+1;row<=range.e.r;row++){
        const codes=codeColumns.map(col=>cellText(sheet[XLSX.utils.encode_cell({r:row,c:range.s.c+col})]));
        if(codes.includes(barcode)){
          const address=XLSX.utils.encode_cell({r:row,c:quantityCol}),previous=sheet[address];
          sheet[address]={t:'n',v:quantity,...(previous?.s?{s:previous.s}:{})};
          found=true;break;
        }
      }
      if(found){matched++;units+=quantity;}else unmatched++;
    }
    const bookType=extension==='csv'?'csv':'biff8';
    output=XLSX.write(workbook,{bookType,type:'base64',sheet:mapping.sheetName,cellStyles:true});
    mimeType=extension==='csv'?'text/csv':'application/vnd.ms-excel';
  }
  const baseName=original.replace(/\.(xlsx|xls|csv)$/i,'');
  return send(res,200,{fileName:`${baseName}-coletado.${extension}`,mimeType,base64:output,matchedCodes:matched,unmatchedCodes:unmatched,totalUnits:units});
 }catch(e){console.error('shared export',e);return send(res,500,{error:'Não foi possível gerar a planilha consolidada'});}
}
