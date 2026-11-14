import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {FileBlob,SpreadsheetFile} from 'file:///C:/Users/Khriz%20Coronel/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const path=fileURLToPath(new URL('./Horario_clases.xlsx',import.meta.url));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const s=wb.worksheets.getItem('Horario');
const colors={'Clases':'#DFE7EF','Desarrollo del sistema':'#F5D5AC','Configuraci\u00f3n de equipos':'#CFD7F7','Almuerzo y transporte':'#D9EFD0'};
const counts={};
for(let r=5;r<=14;r++)for(let c=66;c<=70;c++){
 const cell=s.getRange(String.fromCharCode(c)+r);
 let v=cell.values[0][0];
 if(!Object.hasOwn(colors,v)){v='Clases';cell.values=[[v]];}
 cell.format.fill=colors[v];counts[v]=(counts[v]||0)+1;
}
if(counts.Clases!==22||counts['Almuerzo y transporte']!==7||counts['Desarrollo del sistema']!==8||counts['Configuraci\u00f3n de equipos']!==13)throw new Error('Conteo incorrecto');
s.getRange('A1').values=[['HORARIO SEMANAL']];
wb.recalculate();
console.log(JSON.stringify(counts));
console.log((await wb.inspect({kind:'table',range:'Horario!B5:F14',include:'values',tableMaxRows:10,tableMaxCols:5,maxChars:1000})).ndjson);
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);
const img=await wb.render({sheetName:'Horario',range:'A1:F18',scale:1,format:'png'});await fs.writeFile(new URL('./preview_simplificado.png',import.meta.url),new Uint8Array(await img.arrayBuffer()));
