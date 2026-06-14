import ExcelJS from 'exceljs';
import path from 'path';
import yaml from 'js-yaml';
import fs from 'fs';

const workbookPath = path.join(__dirname, '..', 'input', 'RADET.xlsx');
const headersConfig = yaml.load(
  fs.readFileSync(path.join(__dirname, '..', 'config', 'sheetHeaders.yaml'), 'utf8'),
) as Record<string, string[]>;

// Columns to sample distinct values from
const TARGETS: Record<string, string[]> = {
  CombinedRADET: ['TBStatus', 'CurrentARTStatus', 'CareEntryPoint', 'ARTEnrollmentSetting'],
  CombinedHTS:   ['FinalHIVTestResult'],
};

(async () => {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(workbookPath, {
    sharedStrings: 'cache', styles: 'ignore', hyperlinks: 'ignore', worksheets: 'emit',
  });
  reader.read();

  for await (const ws of reader) {
    const sheetName: string = (ws as any).name;
    const targetCols = TARGETS[sheetName];
    if (!targetCols) { for await (const _ of ws) { /* drain */ } continue; }

    const predefined = headersConfig[sheetName];
    const headers: string[] = predefined ? ['', ...predefined] : [];
    let isFirstRow = !predefined;
    const distinct: Record<string, Set<string>> = {};
    for (const col of targetCols) distinct[col] = new Set();

    for await (const row of ws) {
      if (isFirstRow) {
        isFirstRow = false;
        (row as ExcelJS.Row).eachCell((cell, c) => { headers[c] = String(cell.value ?? '').trim(); });
        continue;
      }
      (row as ExcelJS.Row).eachCell({ includeEmpty: false }, (cell, c) => {
        const h = headers[c];
        if (h && distinct[h]) {
          const v = String(cell.value ?? '').trim();
          if (v) distinct[h].add(v);
        }
      });
    }

    console.log(`\n=== ${sheetName} ===`);
    for (const col of targetCols) {
      const vals = [...distinct[col]].sort();
      console.log(`\n  ${col} (${vals.length} distinct):`);
      vals.slice(0, 30).forEach(v => console.log(`    "${v}"`));
      if (vals.length > 30) console.log(`    ... (${vals.length - 30} more)`);
    }
  }
})();
