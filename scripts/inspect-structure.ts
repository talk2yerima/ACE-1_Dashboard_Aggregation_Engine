import ExcelJS from 'exceljs';
import path from 'path';

const workbookPath = path.join(__dirname, '..', 'input', 'RADET.xlsx');
const TARGET_SHEET = 'CombinedRADET';
const SAMPLE_ROWS = 5;
const MAX_COLS = 120;

(async () => {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(workbookPath, {
    sharedStrings: 'cache', styles: 'ignore', hyperlinks: 'ignore', worksheets: 'emit',
  });
  reader.read();

  for await (const ws of reader) {
    const name: string = (ws as any).name;

    if (name !== TARGET_SHEET) {
      for await (const _ of ws) { /* drain */ }
      continue;
    }

    let rowNum = 0;
    // colIndex -> { samples, looksLikeDatim, looksLikeText }
    const colProfiles: Map<number, { samples: string[]; hasDatim: boolean; hasText: boolean }> = new Map();

    for await (const row of ws) {
      rowNum++;
      if (rowNum > SAMPLE_ROWS) break;

      (row as ExcelJS.Row).eachCell({ includeEmpty: false }, (cell, col) => {
        if (col > MAX_COLS) return;
        const val = String(cell.value ?? '').trim();
        if (!colProfiles.has(col)) colProfiles.set(col, { samples: [], hasDatim: false, hasText: false });
        const p = colProfiles.get(col)!;
        if (p.samples.length < 3) p.samples.push(val.substring(0, 50));

        // 6-7 digit numeric → likely DATIM/facility code
        if (/^\d{5,7}$/.test(val)) p.hasDatim = true;
        // UUID v4 pattern
        const isUuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(val);
        const isEnrollmentKey = /^[A-Za-z0-9]{11}_/.test(val);
        // readable text: not a UUID, not an enrollment key, not purely numeric, min 3 chars
        if (!isUuid && !isEnrollmentKey && /[A-Za-z]/.test(val) && val.length >= 3) p.hasText = true;
      });
    }

    console.log(`\n=== ${name} — column profile (first ${SAMPLE_ROWS} rows) ===\n`);

    for (const [col, p] of [...colProfiles.entries()].sort((a, b) => a[0] - b[0])) {
      const flags: string[] = [];
      if (p.hasDatim)  flags.push('DATIM-CODE?');
      if (p.hasText)   flags.push('READABLE-TEXT?');
      console.log(`  Col ${String(col).padStart(3)}: [${flags.join(', ').padEnd(30)}]  ${p.samples.join(' | ')}`);
    }
  }
})();
