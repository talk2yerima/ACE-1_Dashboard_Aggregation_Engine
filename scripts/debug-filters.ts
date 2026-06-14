/**
 * debug-filters.ts — prints raw + mapped values for the first 10 rows
 * of the filter columns to confirm whether UID→name mapping is working.
 */
import ExcelJS from 'exceljs';
import path from 'path';
import yaml from 'js-yaml';
import fs from 'fs';

const workbookPath = path.join(__dirname, '..', 'input', 'RADET.xlsx');

const headersConfig = yaml.load(
  fs.readFileSync(path.join(__dirname, '..', 'config', 'sheetHeaders.yaml'), 'utf8'),
) as Record<string, string[]>;

const mappingsCfg = yaml.load(
  fs.readFileSync(path.join(__dirname, '..', 'config', 'mappings.yaml'), 'utf8'),
) as { mappings: Record<string, Record<string, string>> };

// Verify YAML loaded correctly
console.log('=== Mapping sanity check ===');
const artMap = mappingsCfg.mappings?.['CurrentARTStatus'];
console.log('CurrentARTStatus keys:', artMap ? Object.keys(artMap).slice(0, 4) : 'NOT FOUND');
const tbMap = mappingsCfg.mappings?.['TBStatus'];
console.log('TBStatus keys:', tbMap ? Object.keys(tbMap).slice(0, 4) : 'NOT FOUND');
const hivMap = mappingsCfg.mappings?.['FinalHIVTestResult'];
console.log('FinalHIVTestResult keys:', hivMap ? Object.keys(hivMap).slice(0, 4) : 'NOT FOUND');

function mapVal(col: string, raw: unknown): string {
  if (raw === null || raw === undefined) return '';
  const str =
    raw instanceof Object && 'result' in (raw as object)
      ? String((raw as any).result ?? '').trim()
      : String(raw).trim();
  const colMap = mappingsCfg.mappings?.[col];
  return colMap ? (colMap[str] ?? str) : str;
}

(async () => {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(workbookPath, {
    sharedStrings: 'cache', styles: 'ignore', hyperlinks: 'ignore', worksheets: 'emit',
  });
  reader.read();

  for await (const ws of reader) {
    const sheetName: string = (ws as any).name;

    if (sheetName === 'CombinedRADET') {
      const headers: string[] = ['', ...headersConfig['CombinedRADET']];
      let rowNum = 0, printed = 0;

      console.log('\n=== CombinedRADET — first 10 rows: raw + mapped ===');
      for await (const row of ws) {
        rowNum++;
        if (printed >= 10) { for await (const _ of ws) {} break; }
        const r: Record<string, unknown> = {};
        (row as ExcelJS.Row).eachCell({ includeEmpty: false }, (cell, c) => {
          const h = headers[c]; if (h) r[h] = cell.value;
        });

        const rawArt = r['CurrentARTStatus'];
        const rawTb  = r['TBStatus'];
        const artStr = mapVal('CurrentARTStatus', rawArt);
        const tbStr  = mapVal('TBStatus', rawTb);

        console.log(`  Row ${rowNum}: CurrentARTStatus raw="${String(rawArt ?? '').slice(0,60)}" → "${artStr}" | TBStatus raw="${String(rawTb ?? '').slice(0,60)}" → "${tbStr}"`);
        printed++;
      }

    } else if (sheetName === 'CombinedHTS') {
      const headers: string[] = ['', ...headersConfig['CombinedHTS']];
      let rowNum = 0, printed = 0;

      console.log('\n=== CombinedHTS — first 10 rows: raw + mapped ===');
      for await (const row of ws) {
        rowNum++;
        if (printed >= 10) { for await (const _ of ws) {} break; }
        const r: Record<string, unknown> = {};
        (row as ExcelJS.Row).eachCell({ includeEmpty: false }, (cell, c) => {
          const h = headers[c]; if (h) r[h] = cell.value;
        });

        const rawHiv = r['FinalHIVTestResult'];
        const hivStr = mapVal('FinalHIVTestResult', rawHiv);

        console.log(`  Row ${rowNum}: FinalHIVTestResult raw="${String(rawHiv ?? '').slice(0,60)}" → "${hivStr}"`);
        printed++;
      }

    } else {
      for await (const _ of ws) { /* drain */ }
    }
  }
})();
