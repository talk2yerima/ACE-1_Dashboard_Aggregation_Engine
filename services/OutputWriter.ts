import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import dayjs from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';
import { Logger } from 'winston';
import { DashboardRow } from './AggregationEngine';
import { ValidationIssue } from './AggregationEngine';

dayjs.extend(customParseFormat);

// Columns that must never be auto-converted to dates by Excel.
const TEXT_ONLY_COLS = ['AgeBand', 'DATIMCode', 'State'] as const;

const DASHBOARD_COLUMNS = [
  { header: 'Period',         key: 'Period',         width: 18, style: { numFmt: 'dd/mm/yyyy' } },
  { header: 'State',          key: 'State',           width: 18, style: { numFmt: '@' } },
  { header: 'Facility',       key: 'Facility',        width: 30 },
  { header: 'DATIMCode',      key: 'DATIMCode',       width: 18, style: { numFmt: '@' } },
  { header: 'Indicator',      key: 'Indicator',       width: 25 },
  { header: 'Disaggregation', key: 'Disaggregation',  width: 18 },
  { header: 'Category',       key: 'Category',        width: 22 },
  { header: 'Sex',            key: 'Sex',             width: 10 },
  { header: 'AgeBand',        key: 'AgeBand',         width: 12, style: { numFmt: '@' } },
  { header: 'Value',          key: 'Value',           width: 12 },
  { header: 'Numerator',      key: 'Numerator',       width: 14 },
  { header: 'Denominator',    key: 'Denominator',     width: 14 },
  { header: 'Target',         key: 'Target',          width: 12 },
  { header: 'AchievementPct', key: 'AchievementPct',  width: 16 },
];

export interface OutputWriterOptions {
  outputDir: string;
  logger: Logger;
}

export class OutputWriter {
  private outputDir: string;
  private logger: Logger;

  constructor(options: OutputWriterOptions) {
    this.outputDir = options.outputDir;
    this.logger = options.logger;
    fs.mkdirSync(this.outputDir, { recursive: true });
  }

  async writeDashboard(rows: DashboardRow[]): Promise<{ xlsx: string; csv: string }> {
    const xlsxPath = path.join(this.outputDir, 'DashboardSummary.xlsx');
    const csvPath = path.join(this.outputDir, 'DashboardSummary.csv');

    await this.writeXlsx(rows, xlsxPath);
    this.writeCsv(rows, csvPath);

    this.logger.info(`OutputWriter: DashboardSummary.xlsx written (${rows.length} rows)`);
    this.logger.info(`OutputWriter: DashboardSummary.csv written (${rows.length} rows)`);

    return { xlsx: xlsxPath, csv: csvPath };
  }

  // ── Streaming CSV (low-memory path for large runs) ──────────────────────────

  openCsvStream(csvPath: string): fs.WriteStream {
    const stream = fs.createWriteStream(csvPath, { encoding: 'utf8' });
    stream.write(DASHBOARD_COLUMNS.map(c => c.header).join(',') + '\n');
    return stream;
  }

  writeCsvRow(stream: fs.WriteStream, row: DashboardRow): void {
    const periodParsed = dayjs(row.Period, 'DD/MM/YYYY', true);
    const periodStr = periodParsed.isValid() ? periodParsed.format('YYYY-MM-DD') : row.Period;
    const line = [
      this.csvCell(periodStr),
      this.csvCell(row.State),
      this.csvCell(row.Facility),
      this.csvCell(row.DATIMCode),
      this.csvCell(row.Indicator),
      this.csvCell(row.Disaggregation),
      this.csvCell(row.Category),
      this.csvCell(row.Sex),
      this.csvText(row.AgeBand),
      this.csvCell(row.Value),
      this.csvCell(row.Numerator),
      this.csvCell(row.Denominator),
      this.csvCell(row.Target),
      this.csvCell(row.AchievementPct !== null && row.AchievementPct !== undefined ? `${row.AchievementPct}%` : ''),
    ].join(',') + '\n';
    stream.write(line);
  }

  closeCsvStream(stream: fs.WriteStream): Promise<void> {
    return new Promise((resolve, reject) => {
      stream.end();
      stream.once('finish', resolve);
      stream.once('error', reject);
    });
  }

  async writeValidationReport(issues: ValidationIssue[]): Promise<string> {
    const reportPath = path.join(this.outputDir, 'ValidationReport.xlsx');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Validation Issues');

    ws.columns = [
      { header: 'Sheet', key: 'sheet', width: 22 },
      { header: 'Row', key: 'row', width: 8 },
      { header: 'Column', key: 'column', width: 25 },
      { header: 'Severity', key: 'severity', width: 12 },
      { header: 'Message', key: 'message', width: 60 },
    ];

    this.styleHeaderRow(ws, 1);

    for (const issue of issues) {
      const wsRow = ws.addRow(issue);
      if (issue.severity === 'ERROR') {
        wsRow.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
        });
      } else if (issue.severity === 'WARNING') {
        wsRow.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
        });
      }
    }

    // Summary tab
    const sumWs = wb.addWorksheet('Summary');
    const errors = issues.filter((i) => i.severity === 'ERROR').length;
    const warnings = issues.filter((i) => i.severity === 'WARNING').length;
    sumWs.addRow(['Total Issues', issues.length]);
    sumWs.addRow(['Errors', errors]);
    sumWs.addRow(['Warnings', warnings]);

    await wb.xlsx.writeFile(reportPath);
    this.logger.info(`OutputWriter: ValidationReport.xlsx written (${issues.length} issues)`);
    return reportPath;
  }

  private async writeXlsx(rows: DashboardRow[], filePath: string): Promise<void> {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'RADET Dashboard Engine';
    wb.created = new Date();

    const ws = wb.addWorksheet('DashboardSummary');
    ws.columns = DASHBOARD_COLUMNS;
    this.styleHeaderRow(ws, 1);

    for (const row of rows) {
      // Parse the Period string to a real Date when possible (DD/MM/YYYY)
      const periodParsed = dayjs(row.Period, 'DD/MM/YYYY', true);
      const periodCell   = periodParsed.isValid() ? periodParsed.toDate() : row.Period;

      const wsRow = ws.addRow({
        ...row,
        Period: periodCell,
        AchievementPct:
          row.AchievementPct !== null && row.AchievementPct !== undefined
            ? `${row.AchievementPct}%`
            : '',
        Value:       row.Value       ?? 0,
        Numerator:   row.Numerator   ?? '',
        Denominator: row.Denominator ?? '',
        Target:      row.Target      ?? '',
      });

      // Enforce text format on identifier cells to prevent Excel auto-conversion
      for (const col of TEXT_ONLY_COLS) {
        const cell = wsRow.getCell(col);
        cell.value  = String(cell.value ?? '');
        cell.numFmt = '@';
      }
      // Ensure date cells get proper date format
      if (periodParsed.isValid()) {
        wsRow.getCell('Period').numFmt = 'dd/mm/yyyy';
      }
    }

    // Freeze header row
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    // Auto-filter
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: DASHBOARD_COLUMNS.length },
    };

    // Indicator pivot sheet
    await this.addPivotSheet(wb, rows);

    await wb.xlsx.writeFile(filePath);
  }

  private async addPivotSheet(wb: ExcelJS.Workbook, rows: DashboardRow[]): Promise<void> {
    const ws = wb.addWorksheet('IndicatorSummary');
    ws.columns = [
      { header: 'Indicator', key: 'Indicator', width: 30 },
      { header: 'Total Value', key: 'Total', width: 15 },
      { header: 'Facilities', key: 'Facilities', width: 15 },
    ];
    this.styleHeaderRow(ws, 1);

    const byIndicator = new Map<string, { total: number; facilities: Set<string> }>();
    for (const row of rows) {
      const ind = row.Indicator;
      if (!byIndicator.has(ind)) byIndicator.set(ind, { total: 0, facilities: new Set() });
      const entry = byIndicator.get(ind)!;
      entry.total += row.Value ?? 0;
      if (row.Facility) entry.facilities.add(row.Facility);
    }

    for (const [indicator, data] of byIndicator.entries()) {
      ws.addRow({ Indicator: indicator, Total: data.total, Facilities: data.facilities.size });
    }
  }

  private writeCsv(rows: DashboardRow[], filePath: string): void {
    const headers = DASHBOARD_COLUMNS.map((c) => c.header).join(',');
    const lines = [headers];

    for (const row of rows) {
      // Write Period as ISO date (YYYY-MM-DD) so Excel auto-parses it correctly
      const periodParsed = dayjs(row.Period, 'DD/MM/YYYY', true);
      const periodStr = periodParsed.isValid() ? periodParsed.format('YYYY-MM-DD') : row.Period;
      const line = [
        this.csvCell(periodStr),
        this.csvCell(row.State),
        this.csvCell(row.Facility),
        this.csvCell(row.DATIMCode),
        this.csvCell(row.Indicator),
        this.csvCell(row.Disaggregation),
        this.csvCell(row.Category),
        this.csvCell(row.Sex),
        this.csvText(row.AgeBand),     // date-like: protect with ="…"
        this.csvCell(row.Value),
        this.csvCell(row.Numerator),
        this.csvCell(row.Denominator),
        this.csvCell(row.Target),
        this.csvCell(row.AchievementPct !== null && row.AchievementPct !== undefined ? `${row.AchievementPct}%` : ''),
      ].join(',');
      lines.push(line);
    }

    fs.writeFileSync(filePath, lines.join('\n'), 'utf8');
  }

  /** Standard quoted CSV cell — handles commas, quotes, and newlines. */
  private csvCell(val: unknown): string {
    if (val === null || val === undefined) return '""';
    return `"${String(val).replace(/"/g, '""')}"`;
  }

  /**
   * Excel-safe text cell.  Uses the ="…" formula prefix so Excel evaluates
   * the cell as a text string and never auto-converts values like "1-4" to
   * "4-Jan" or "12/06/2026" to a date serial, even when the CSV is opened
   * by double-click.  The formula result is the plain text value; filtering
   * and sorting inside Excel work correctly on the resolved string.
   */
  private csvText(val: unknown): string {
    if (val === null || val === undefined) return '""';
    const str = String(val).replace(/"/g, '""');
    return `"=""${str}"""`;
  }

  private styleHeaderRow(ws: ExcelJS.Worksheet, rowNum: number): void {
    const headerRow = ws.getRow(rowNum);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203864' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
        right: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      };
    });
    headerRow.height = 22;
  }
}
