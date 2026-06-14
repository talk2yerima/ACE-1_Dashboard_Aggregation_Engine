import ExcelJS from 'exceljs';
import path from 'path';

const workbookPath = path.join(__dirname, '..', 'input', 'RADET.xlsx');

(async () => {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(workbookPath, {
    sharedStrings: 'cache', styles: 'ignore', hyperlinks: 'ignore', worksheets: 'emit',
  });
  reader.read();

  for await (const ws of reader) {
    const name: string = (ws as any).name;
    let done = false;
    for await (const row of ws) {
      if (!done) {
        const headers: string[] = [];
        (row as ExcelJS.Row).eachCell({ includeEmpty: false }, (cell, col) => {
          headers[col] = String(cell.value ?? '').trim();
        });
        console.log(`\n=== ${name} ===`);
        headers.filter(Boolean).forEach((h, i) => console.log(`  [${i + 1}] ${h}`));
        done = true;
      }
      // drain remaining rows
    }
  }
})();
