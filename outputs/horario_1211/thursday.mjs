import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {FileBlob,SpreadsheetFile} from 'file:///C:/Users/Khriz%20Coronel/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const path=fileURLToPath(new URL('./Horario_clases.xlsx',import.meta.url));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));const s=wb.worksheets.getItem('Horario');
const before=s.getRange('B5:F14').values;
s.getRange('E10:E11').values=[['Configuraci\u00f3n de equipos']];
s.getRange('E10:E11').format.fill='#CFD7F7';
s.getRange('A16').values=[['22 h de clases | 8 h de desarrollo | 15 h de configuraci\u00f3n | 5 h de almuerzo y transporte']];
s.getRange('A18').values=[['Jueves: almuerzo y transporte de 11:30 a 12:30; configuraci\u00f3n de equipos de 12:30 a 17:30.']];
const after=s.getRange('B5:F14').values,counts={};let changed=0;
for(let r=0;r<10;r++)for(let c=0;c<5;c++){const v=after[r][c];counts[v]=(counts[v]||0)+1;if(v!==before[r][c]){changed++;if(c!==3||![5,6].includes(r))throw new Error('Cambio fuera de alcance');}}
if(changed!==2||counts['Configuraci\u00f3n de equipos']!==15||counts['Almuerzo y transporte']!==5)throw new Error('Verificaci\u00f3n fallida');
wb.recalculate();console.log(JSON.stringify(counts));
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(path);
const img=await wb.render({sheetName:'Horario',range:'A1:F18',scale:1,format:'png'});await fs.writeFile(new URL('./preview_jueves.png',import.meta.url),new Uint8Array(await img.arrayBuffer()));
