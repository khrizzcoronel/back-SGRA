import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {FileBlob,SpreadsheetFile} from 'file:///C:/Users/Khriz%20Coronel/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const path=fileURLToPath(new URL('./Horario_clases.xlsx',import.meta.url));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));const s=wb.worksheets.getItem('Horario');
let dev=0,clases=0,almuerzo=0;
for(let r=5;r<=14;r++)for(let c=66;c<=70;c++){
 const cell=s.getRange(String.fromCharCode(c)+r),v=cell.values[0][0];
 if(v==='Clases'){clases++;continue;}
 if(v==='Almuerzo y transporte'){almuerzo++;continue;}
 if(!['Desarrollo del sistema','Configuraci\u00f3n de equipos'].includes(v))throw new Error('Actividad inesperada');
 cell.values=[['Desarrollo']];cell.format.fill='#F5D5AC';dev++;
}
if(dev!==23||clases!==22||almuerzo!==5)throw new Error('Totales incorrectos');
s.getRange('A1').values=[['HORARIO SEMANAL — DESARROLLO']];
s.getRange('A16').values=[['22 h de clases | 23 h de desarrollo | 5 h de almuerzo y transporte']];
s.getRange('A18').values=[['Jueves: almuerzo y transporte de 11:30 a 12:30; desarrollo de 12:30 a 17:30.']];
wb.recalculate();console.log(JSON.stringify({desarrollo:dev,clases,almuerzo}));
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(fileURLToPath(new URL('./Horario_desarrollo.xlsx',import.meta.url)));
const img=await wb.render({sheetName:'Horario',range:'A1:F18',scale:1,format:'png'});await fs.writeFile(new URL('./preview_desarrollo.png',import.meta.url),new Uint8Array(await img.arrayBuffer()));
