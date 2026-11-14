import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from 'file:///C:/Users/Khriz%20Coronel/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const wb=Workbook.create(),s=wb.worksheets.add('Horario');
const a='Aplicaciones distribuidas',g='Seguridad informática',i='Inteligencia de negocios',m='Mantenimiento de software',p='Proceso de software',l='Libre';
const data=[[a,g,i,a,m],[a,g,i,a,m],[a,i,l,g,l],[m,i,l,g,l],[m,i,l,l,l],[l,l,l,l,p],[l,l,p,l,p],[l,l,p,l,l],[l,l,l,l,l],[l,l,l,l,l]];
s.getRange('A1:F16').format.font={name:'Aptos',size:11,color:'#243247'};
s.getRange('A1:F1').merge();s.getRange('A1').values=[['HORARIO DE CLASES']];s.getRange('A1:F1').format={fill:'#17324D',font:{bold:true,size:22,color:'#FFFFFF'},rowHeight:42};
s.getRange('A2:F2').merge();s.getRange('A2').values=[['Lunes a viernes | 07:30–17:30 | 26/10/2026–28/02/2027']];s.getRange('A2:F2').format.rowHeight=28;
s.getRange('A4:F4').values=[['Hora','Lunes','Martes','Miércoles','Jueves','Viernes']];s.getRange('A4:F4').format={fill:'#17324D',font:{bold:true,color:'#FFFFFF'},rowHeight:30};
const pad=n=>String(n).padStart(2,'0');s.getRange('A5:F14').values=data.map((d,k)=>[pad(7+k)+':30 – '+pad(8+k)+':30',...d]);
s.getRange('A4:F14').format.horizontalAlignment='center';s.getRange('A4:F14').format.verticalAlignment='center';s.getRange('A5:F14').format.rowHeight=54;s.getRange('B5:F14').format.wrapText=true;s.getRange('A4:A14').format.columnWidth=23;s.getRange('B4:F14').format.columnWidth=28;s.getRange('A5:A14').format.fill='#E8EEF4';s.getRange('A5:A14').format.font.bold=true;s.getRange('A4:F14').format.borders={preset:'all',style:'thin',color:'#FFFFFF'};
const colors={[a]:'#D6E9FA',[g]:'#E4DCF5',[i]:'#D8EFE2',[m]:'#FFE8C9',[p]:'#CFEFF0',[l]:'#F0F2F5'};
data.forEach((row,r)=>row.forEach((v,c)=>{const cell=s.getRange(String.fromCharCode(66+c)+(r+5));cell.format.fill=colors[v];cell.format.font.bold=v!==l;cell.format.font.color=v===l?'#77828F':'#243247';}));
s.getRange('A16:F16').merge();s.getRange('A16').values=[['24 horas de clase semanales · Los espacios sin clases están marcados como Libre.']];s.getRange('A16:F16').format.rowHeight=25;
wb.recalculate();console.log((await wb.inspect({kind:'table',range:'Horario!A4:F14',include:'values',tableMaxRows:11,tableMaxCols:6,maxChars:1700})).ndjson);
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(decodeURI(new URL('./Horario_clases.xlsx',import.meta.url).pathname).replace(/^\/(\w:)/,'$1'));
const img=await wb.render({sheetName:'Horario',range:'A1:F16',scale:1,format:'png'});await fs.writeFile(new URL('./preview.png',import.meta.url),new Uint8Array(await img.arrayBuffer()));

