/**
 * all-status-values.ts
 * Frequency scan of ALL dimension/filter columns so we can identify
 * every DHIS2 UID → human-readable name mapping needed.
 */
import ExcelJS from 'exceljs';
import path from 'path';
import yaml from 'js-yaml';
import fs from 'fs';

const workbookPath = path.join(__dirname, '..', 'input', 'RADET.xlsx');
const headersConfig = yaml.load(
  fs.readFileSync(path.join(__dirname, '..', 'config', 'sheetHeaders.yaml'), 'utf8'),
) as Record<string, string[]>;

// Columns to scan per sheet
const TARGETS: Record<string, string[]> = {
  CombinedRADET: [
    'State', 'LGA', 'Facility', 'DATIMCode',
    'Sex', 'CareEntryPoint', 'ARTEnrollmentSetting',
    'CurrentARTStatus', 'TBStatus',
  ],
  CombinedHTS: ['Sex', 'FinalHIVTestResult'],
};

(async () => {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(workbookPath, {
    sharedStrings: 'cache', styles: 'ignore', hyperlinks: 'ignore', worksheets: 'emit',
  });
  reader.read();

  for await (const ws of reader) {
    const sheetName: string = (ws as any).name;
    const targetCols = TARGETS[sheetName];
    if (!targetCols) { for await (const _ of ws) {} continue; }

    const predefined = headersConfig[sheetName];
    const headers: string[] = predefined ? ['', ...predefined] : [];
    const freq: Record<string, Map<string, number>> = {};
    for (const col of targetCols) freq[col] = new Map();

    for await (const row of ws) {
      (row as ExcelJS.Row).eachCell({ includeEmpty: false }, (cell, c) => {
        const h = headers[c];
        if (h && freq[h]) {
          const v = String(cell.value ?? '').trim();
          if (v) freq[h].set(v, (freq[h].get(v) ?? 0) + 1);
        }
      });
    }

    console.log(`\n${'='.repeat(64)}\nSheet: ${sheetName}\n${'='.repeat(64)}`);
    for (const col of targetCols) {
      const entries = [...freq[col].entries()].sort((a, b) => b[1] - a[1]);
      const total   = entries.reduce((s, [, n]) => s + n, 0);
      console.log(`\n  [${col}]  ${entries.length} distinct, ${total} non-empty`);
      // Print top 20 by frequency
      for (const [v, n] of entries.slice(0, 20)) {
        const pct = ((n / total) * 100).toFixed(1).padStart(5);
        console.log(`    ${n.toString().padStart(7)}  ${pct}%   "${v}"`);
      }
      if (entries.length > 20) console.log(`    ... (${entries.length - 20} more)`);
    }
  }
})();
