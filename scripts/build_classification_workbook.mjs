// Render the consolidated classification tables without changing mapping values.
// Usage: node scripts/build_classification_workbook.mjs tables.json output.xlsx module-root
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';

const [input, output, moduleRoot] = process.argv.slice(2);
if (!input || !output || !moduleRoot) throw new Error('Supply tables.json, output.xlsx and the bundled node_modules directory.');
const require = createRequire(path.join(path.resolve(moduleRoot), 'classification-runtime.cjs'));
const {Workbook, SpreadsheetFile} = await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const data = JSON.parse(await fs.readFile(input, 'utf8'));
const wb = Workbook.create();
const previews = path.join(path.dirname(output), 'previews');
await fs.mkdir(previews, {recursive:true});
const sources = data.sheets.find(s => s.name === 'Sources');
for (const s of data.sheets) if (s.note) sources.rows.push([s.name, 'Workbook interpretation', s.note]);
for (const [i, definition] of data.sheets.entries()) {
  const s = wb.worksheets.add(definition.name);
  s.showGridLines = false;
  const values = [definition.headers, ...definition.rows];
  const n = definition.headers.length;
  const all = s.getRangeByIndexes(0,0,values.length,n);
  all.values = values;
  all.format.font = {name:'Arial',size:10,color:'#243746'};
  all.format.verticalAlignment = 'center';
  all.format.rowHeight = 28;
  all.setNumberFormat('@');
  definition.widths.forEach((width,c) => {
    const column = s.getRangeByIndexes(0,c,values.length,1);
    column.format.columnWidth = width;
    column.format.wrapText = true;
  });
  // Fit long source descriptions and provenance without shrinking the text.
  values.slice(1).forEach((row,r) => {
    const lines = Math.max(...row.map((v,c) => Math.ceil(String(v ?? '').length / (definition.widths[c] * 1.05))));
    s.getRangeByIndexes(r+1,0,1,n).format.rowHeight = Math.max(28,lines*13+8);
  });
  const header = s.getRangeByIndexes(0,0,1,n);
  header.format.fill = '#23394D';
  header.format.font = {name:'Arial',size:10,bold:true,color:'#FFFFFF'};
  header.format.rowHeight = 40;
  header.format.horizontalAlignment = 'center';
  s.tables.add(all,true,`ClassificationTable${i+1}`);
  s.freezePanes.freezeRows(1);
  if (n > 3) s.freezePanes.freezeColumns(2);
  if (definition.name === 'ENTICE sectors') {
    s.tabColor = '#23394D';
    for (const text of ['unreviewed','scope_required','bridge_missing']) {
      s.getRange(`F2:F${values.length}`).conditionalFormats.add('containsText',{
        text,format:{fill:'#FFF2CC',font:{color:'#7F6000'}}
      });
    }
  }
}
wb.recalculate();
const checks = [];
for (const [i,s] of data.sheets.entries()) {
  const result = await wb.inspect({kind:'table',range:`'${s.name}'!A1:F6`,include:'values,formulas',tableMaxRows:6,tableMaxCols:6,maxChars:2000});
  checks.push(result.ndjson);
  const last = String.fromCharCode(64+Math.min(s.headers.length,6));
  const image = await wb.render({sheetName:s.name,range:`A1:${last}8`,scale:1,format:'png'});
  await fs.writeFile(path.join(previews,`${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await image.arrayBuffer()));
}
const extra = await wb.render({sheetName:'ENTICE sectors',range:'G1:L8',scale:1,format:'png'});
await fs.writeFile(path.join(previews,'01-details.png'),new Uint8Array(await extra.arrayBuffer()));
checks.push((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:30},summary:'Classification export error scan'})).ndjson);
await fs.writeFile(path.join(path.dirname(output),'classification-checks.jsonl'),checks.join('\n'));
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(JSON.stringify({output,sheets:data.sheets.map(s=>({name:s.name,rows:s.rows.length}))}));
