/**
 * freq-values.ts
 * Counts occurrences of each distinct value in the filter columns so we can
 * identify which DHIS2 UIDs correspond to human-readable option-set labels
 * (e.g. the most common FinalHIVTestResult UID is likely "Negative").
 */
import ExcelJS from 'exceljs';
import path from 'path';
import yaml from 'js-yaml';
import fs from 'fs';

const workbookPath = path.join(__dirname, '..', 'input', 'RADET.xlsx');
const headersConfig = yaml.load(
  fs.readFileSync(path.join(__dirname, '..', 'config', 'sheetHeaders.yaml'), 'utf8'),
) as Record<string, string[]>;

// Columns whose value frequencies we want to count
const TARGETS: Record<string, string[]> = {
  CombinedRADET: ['CurrentARTStatus', 'TBStatus', 'CareEntryPoint', 'ARTEnrollmentSetting'],
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
    const freq: Record<string, Map<string, number>> = {};
    for (const col of targetCols) freq[col] = new Map();

    for await (const row of ws) {
      if (isFirstRow) {
        isFirstRow = false;
        (row as ExcelJS.Row).eachCell((cell, c) => { headers[c] = String(cell.value ?? '').trim(); });
        continue;
      }
      (row as ExcelJS.Row).eachCell({ includeEmpty: false }, (cell, c) => {
        const h = headers[c];
        if (h && freq[h]) {
          const v = String(cell.value ?? '').trim();
          if (v) freq[h].set(v, (freq[h].get(v) ?? 0) + 1);
        }
      });
    }

    console.log(`\n=== ${sheetName} ===`);
    for (const col of targetCols) {
      const sorted = [...freq[col].entries()].sort((a, b) => b[1] - a[1]);
      const total = sorted.reduce((s, [, n]) => s + n, 0);
      console.log(`\n  ${col} (${sorted.length} values, ${total} total non-empty):`);
      for (const [v, n] of sorted) {
        const pct = ((n / total) * 100).toFixed(1);
        console.log(`    ${n.toString().padStart(7)}  (${pct.padStart(5)}%)  "${v}"`);
      }
    }
  }
})();
