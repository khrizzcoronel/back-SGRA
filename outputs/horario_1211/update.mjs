import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from 'file:///C:/Users/Khriz%20Coronel/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
import {fileURLToPath} from 'node:url';
const path=fileURLToPath(new URL('./Horario_clases.xlsx',import.meta.url));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const s=wb.worksheets.getItem('Horario');
const d='Desarrollo del sistema',c='Configuraci\u00f3n de equipos',l='Almuerzo y transporte';
const slots={B:[[10,10,l],[11,14,d]],C:[[10,10,l],[11,14,d]],D:[[7,9,c],[10,10,l],[13,14,c]],E:[[9,11,l],[12,14,c]],F:[[7,8,c],[9,9,l],[12,14,c]]};
let dev=0,config=0,breaks=0;
for(const [col,blocks] of Object.entries(slots))for(const [start,end,label] of blocks){
 for(let r=start;r<=end;r++){const cell=s.getRange(col+r);if(![d,c,l].includes(cell.values[0][0]))throw new Error('Conflicto con clase '+col+r);cell.values=[[label]];label===d?dev++:label===c?config++:breaks++;}
 s.getRange(`${col}${start}:${col}${end}`).format={fill:label===d?'#F5D5AC':label===c?'#CFD7F7':'#E7ECEF',font:{bold:true,color:'#243247'},wrapText:true};
}
if(dev!==8||config!==13||breaks!==7)throw new Error('Totales incorrectos');
s.getRange('A1').values=[['HORARIO DE CLASES Y TRABAJO']];
s.getRange('A16').values=[['22 h de clases | 8 h de desarrollo | 13 h de configuraci\u00f3n | 7 h de almuerzo y transporte']];
s.getRange('A18:F18').merge();s.getRange('A18').values=[['Jueves: almuerzo y transporte de 11:30 a 14:30 (3 horas); configuraci\u00f3n desde las 14:30.']];s.getRange('A18:F18').format={font:{name:'Aptos',size:10,color:'#516275'},rowHeight:26};
wb.recalculate();
console.log(JSON.stringify({desarrollo:dev,configuracion:config,almuerzoTransporte:breaks,horas:dev+config+breaks}));
console.log((await wb.inspect({kind:'table',range:'Horario!A4:F14',include:'values',tableMaxRows:11,tableMaxCols:6,maxChars:2200})).ndjson);
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);
const img=await wb.render({sheetName:'Horario',range:'A1:F18',scale:1,format:'png'});await fs.writeFile(new URL('./preview_actualizado.png',import.meta.url),new Uint8Array(await img.arrayBuffer()));
