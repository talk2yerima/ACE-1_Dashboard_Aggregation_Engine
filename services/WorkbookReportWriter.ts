import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import readline from 'readline';
import dayjs, { Dayjs } from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';
import isBetween from 'dayjs/plugin/isBetween';
import { Logger } from 'winston';
import { DashboardRow } from './AggregationEngine';

dayjs.extend(customParseFormat);
dayjs.extend(isBetween);

const EXCEL_MAX_ROWS = 1_048_576;
const FISCAL_QUARTERS = [
  { label: 'Qtr1', start: dayjs('2025-10-01'), end: dayjs('2025-12-31') },
  { label: 'Qtr2', start: dayjs('2026-01-01'), end: dayjs('2026-03-31') },
  { label: 'Qtr3', start: dayjs('2026-04-01'), end: dayjs('2026-06-30') },
  { label: 'Qtr4', start: dayjs('2026-07-01'), end: dayjs('2026-09-30') },
];

// ── EDCT aggregation types ────────────────────────────────────────────────────

/** One aggregated EDCT bucket: a unique Facility × Period × Indicator × Sex × AgeBand */
interface EdctBucket {
  state: string;
  facility: string;
  datim: string;
  period: string;          // quarter label ("Qtr1") or month label ("May-26")
  indicator: string;
  disaggregation: string;
  category: string;
  sex: string;
  ageBand: string;
  value: number;
}

type TargetRow = {
  date: Dayjs;
  values: unknown[];
};

type FacilityRows = {
  state: string;
  lga: string;
  facility: string;
  rows: TargetRow[];
};

export class WorkbookReportWriter {
  constructor(private logger: Logger) {}

  async writeFyDashboardWorkbook(options: {
    outputDir: string;
    dashboardCsvPath: string;
    targetWorkbookPath: string;
    edctRows?: DashboardRow[];
  }): Promise<string> {
    const outputPath = path.join(options.outputDir, 'DashboardSummary.xlsx');
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: outputPath,
      useStyles: true,
      useSharedStrings: false,
    });
    workbook.creator = 'RADET Dashboard Engine';
    workbook.created = new Date();

    await this.addDashboardSheet(workbook, options.dashboardCsvPath);
    await this.addTargetPeriodSheets(workbook, options.targetWorkbookPath, options.edctRows ?? []);

    await workbook.commit();
    this.logger.info(`OutputWriter: DashboardSummary.xlsx written -> ${outputPath}`);
    return outputPath;
  }

  // ── Sheet 1: Dashboard (all daily rows from DashboardSummary.csv) ────────────

  private async addDashboardSheet(workbook: ExcelJS.stream.xlsx.WorkbookWriter, csvPath: string): Promise<void> {
    const ws = workbook.addWorksheet('Dashboard');
    let rowCount = 0;
    let headers: string[] = [];

    const rl = readline.createInterface({
      input: fs.createReadStream(csvPath, { encoding: 'utf8' }),
      crlfDelay: Infinity,
    });

    for await (const line of rl) {
      if (!line.trim()) continue;
      const parsed = this.parseCsvLine(line);
      if (rowCount === 0) {
        headers = parsed;
        ws.columns = headers.map((h) => ({ header: h, key: h, width: this.widthForHeader(h) }));
        this.styleHeaderRow(ws.getRow(1));
        rowCount++;
        continue;
      }
      if (rowCount >= EXCEL_MAX_ROWS) {
        this.logger.warn(`Dashboard sheet reached Excel row limit; remaining CSV rows stayed in ${csvPath}`);
        break;
      }
      ws.addRow(this.coerceDashboardRow(parsed, headers)).commit();
      rowCount++;
    }
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    ws.commit();
  }

  // ── Sheets 2 & 3: Quarterly / Monthly (targets + EDCT aggregated) ────────────

  private async addTargetPeriodSheets(
    workbook: ExcelJS.stream.xlsx.WorkbookWriter,
    targetWorkbookPath: string,
    edctRows: DashboardRow[],
  ): Promise<void> {
    const source = new ExcelJS.Workbook();
    await source.xlsx.readFile(targetWorkbookPath);
    const sheet = source.worksheets[0];
    if (!sheet) throw new Error(`No sheet found in ${targetWorkbookPath}`);

    const headers = sheet.getRow(1).values as unknown[];
    const cleanHeaders = headers.slice(1).map((h) => String(h ?? '').trim());
    const stateCol    = cleanHeaders.indexOf('STATE')    + 1;
    const lgaCol      = cleanHeaders.indexOf('LGA')      + 1;
    const facilityCol = cleanHeaders.indexOf('FACILITY') + 1;
    const dateCol     = cleanHeaders.indexOf('DATE')     + 1;
    const metricHeaders  = cleanHeaders.filter((h) => !['STATE', 'LGA', 'FACILITY', 'Qtr', 'DATE'].includes(h));
    const metricIndexes  = metricHeaders.map((h) => cleanHeaders.indexOf(h) + 1);

    // Collect unique EDCT indicator names (preserves insertion order = SQL order)
    const edctIndicators = this.uniqueEdctIndicators(edctRows);
    this.logger.info(`  EDCT indicators for quarterly/monthly sheets: ${edctIndicators.join(', ') || '(none)'}`);

    const facilities  = new Map<string, FacilityRows>();
    const monthStarts = new Map<string, Dayjs>();
    let fyMaxDate: Dayjs | null = null;

    for (let r = 2; r <= sheet.rowCount; r++) {
      const row  = sheet.getRow(r);
      const date = this.parseExcelDate(row.getCell(dateCol).value);
      if (!date) continue;
      const state    = String(row.getCell(stateCol).value    ?? '').trim();
      const lga      = String(row.getCell(lgaCol).value      ?? '').trim();
      const facility = String(row.getCell(facilityCol).value ?? '').trim();
      const key = `${state.toLowerCase()}|${this.normalizeFacility(facility)}`;
      if (!facilities.has(key)) facilities.set(key, { state, lga, facility: this.stripStatePrefix(facility), rows: [] });
      facilities.get(key)!.rows.push({ date, values: metricIndexes.map((idx) => row.getCell(idx).value ?? null) });
      monthStarts.set(date.format('YYYY-MM'), date.startOf('month'));
      if (!fyMaxDate || date.isAfter(fyMaxDate)) fyMaxDate = date;
    }

    for (const item of facilities.values()) {
      item.rows.sort((a, b) => a.date.valueOf() - b.date.valueOf());
    }

    const months = [...monthStarts.values()].sort((a, b) => a.valueOf() - b.valueOf());

    // Aggregate EDCT by quarter and month
    const edctQuarterly = this.aggregateEdctByQuarter(edctRows);
    const edctMonthly   = this.aggregateEdctByMonth(edctRows);

    this.writeQuarterlySheet(workbook, [...facilities.values()], metricHeaders, fyMaxDate, edctQuarterly, edctIndicators);
    this.writeMonthlySheet(workbook, [...facilities.values()], months, metricHeaders, edctMonthly, edctIndicators);
  }

  // ── EDCT aggregation ─────────────────────────────────────────────────────────

  private uniqueEdctIndicators(rows: DashboardRow[]): string[] {
    const seen = new Set<string>();
    const result: string[] = [];
    for (const r of rows) {
      if (!seen.has(r.Indicator)) { seen.add(r.Indicator); result.push(r.Indicator); }
    }
    return result;
  }

  /** Group EDCT rows by (State, Facility, DATIMCode, Indicator, Disaggregation, Category, Sex, AgeBand, Quarter) → SUM Value */
  private aggregateEdctByQuarter(rows: DashboardRow[]): EdctBucket[] {
    const map = new Map<string, EdctBucket>();

    for (const row of rows) {
      const period = dayjs(row.Period, 'DD/MM/YYYY', true);
      if (!period.isValid()) continue;

      // FY26 cumulative in addition to individual quarters
      const quarters = FISCAL_QUARTERS.filter((q) =>
        (period.isSame(q.start, 'day') || period.isAfter(q.start)) &&
        (period.isSame(q.end,   'day') || period.isBefore(q.end)),
      );
      const fy26Start = dayjs('2025-10-01');
      const fy26End   = dayjs('2026-09-30');
      const inFy26 =
        (period.isSame(fy26Start, 'day') || period.isAfter(fy26Start)) &&
        (period.isSame(fy26End,   'day') || period.isBefore(fy26End));
      if (inFy26) quarters.push({ label: 'FY26', start: fy26Start, end: fy26End });

      for (const q of quarters) {
        const key = [row.DATIMCode, q.label, row.Indicator, row.Disaggregation, row.Category, row.Sex, row.AgeBand].join('|');
        if (map.has(key)) {
          map.get(key)!.value += row.Value ?? 0;
        } else {
          map.set(key, {
            state:         row.State,
            facility:      row.Facility,
            datim:         row.DATIMCode,
            period:        q.label,
            indicator:     row.Indicator,
            disaggregation: row.Disaggregation,
            category:      row.Category,
            sex:           row.Sex,
            ageBand:       row.AgeBand,
            value:         row.Value ?? 0,
          });
        }
      }
    }
    return [...map.values()];
  }

  /** Group EDCT rows by (State, Facility, DATIMCode, Indicator, Disaggregation, Category, Sex, AgeBand, Month) → SUM Value */
  private aggregateEdctByMonth(rows: DashboardRow[]): EdctBucket[] {
    const map = new Map<string, EdctBucket>();

    for (const row of rows) {
      const period = dayjs(row.Period, 'DD/MM/YYYY', true);
      if (!period.isValid()) continue;
      const monthLabel = period.format('MMM-YY');   // e.g. "May-26"

      const key = [row.DATIMCode, monthLabel, row.Indicator, row.Disaggregation, row.Category, row.Sex, row.AgeBand].join('|');
      if (map.has(key)) {
        map.get(key)!.value += row.Value ?? 0;
      } else {
        map.set(key, {
          state:         row.State,
          facility:      row.Facility,
          datim:         row.DATIMCode,
          period:        monthLabel,
          indicator:     row.Indicator,
          disaggregation: row.Disaggregation,
          category:      row.Category,
          sex:           row.Sex,
          ageBand:       row.AgeBand,
          value:         row.Value ?? 0,
        });
      }
    }
    return [...map.values()];
  }

  // ── Sheet writers ─────────────────────────────────────────────────────────────

  private writeQuarterlySheet(
    workbook: ExcelJS.stream.xlsx.WorkbookWriter,
    facilities: FacilityRows[],
    metricHeaders: string[],
    fyMaxDate: Dayjs | null,
    edctBuckets: EdctBucket[],
    edctIndicators: string[],
  ): void {
    const ws = workbook.addWorksheet('Quarterly');

    // Shared header: target columns + EDCT indicator columns
    const headers = ['STATE', 'LGA', 'FACILITY', 'Quarter', 'TargetDate', ...metricHeaders, ...edctIndicators];
    ws.columns = headers.map((h) => ({ header: h, key: h, width: this.widthForHeader(h) }));
    this.styleHeaderRow(ws.getRow(1));

    const nullEdct = edctIndicators.map(() => null);

    // ── existing target rows (EDCT columns blank) ──
    for (const facility of facilities) {
      for (const quarter of FISCAL_QUARTERS) {
        const row = this.findMaxDateRow(facility.rows, quarter.start, quarter.end);
        if (!row) continue;
        ws.addRow([
          facility.state, facility.lga, facility.facility,
          quarter.label, row.date.format('YYYY-MM-DD'),
          ...row.values,
          ...nullEdct,
        ]).commit();
      }
      if (fyMaxDate) {
        const row = this.findMaxDateRow(facility.rows, dayjs('2025-10-01'), fyMaxDate);
        if (row) ws.addRow([
          facility.state, facility.lga, facility.facility,
          'FY26', row.date.format('YYYY-MM-DD'),
          ...row.values,
          ...nullEdct,
        ]).commit();
      }
    }

    // ── EDCT rows (target metric columns blank) ──
    const nullMetrics = metricHeaders.map(() => null);
    for (const bucket of edctBuckets) {
      const edctValues = edctIndicators.map((ind) => (ind === bucket.indicator ? bucket.value : null));
      ws.addRow([
        bucket.state, '', bucket.facility,
        bucket.period, '',
        ...nullMetrics,
        ...edctValues,
      ]).commit();
    }

    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    ws.commit();
    this.logger.info(`  Quarterly sheet: ${facilities.length} target facilities + ${edctBuckets.length} EDCT rows`);
  }

  private writeMonthlySheet(
    workbook: ExcelJS.stream.xlsx.WorkbookWriter,
    facilities: FacilityRows[],
    months: Dayjs[],
    metricHeaders: string[],
    edctBuckets: EdctBucket[],
    edctIndicators: string[],
  ): void {
    const ws = workbook.addWorksheet('Monthly');

    const headers = ['STATE', 'LGA', 'FACILITY', 'Month', 'TargetDate', ...metricHeaders, ...edctIndicators];
    ws.columns = headers.map((h) => ({ header: h, key: h, width: this.widthForHeader(h) }));
    this.styleHeaderRow(ws.getRow(1));

    const nullEdct = edctIndicators.map(() => null);

    // ── existing target rows ──
    for (const facility of facilities) {
      for (const month of months) {
        const row = this.findMaxDateRow(facility.rows, month.startOf('month'), month.endOf('month'));
        if (!row) continue;
        ws.addRow([
          facility.state, facility.lga, facility.facility,
          month.format('MMM-YY'), row.date.format('YYYY-MM-DD'),
          ...row.values,
          ...nullEdct,
        ]).commit();
      }
    }

    // ── EDCT rows ──
    const nullMetrics = metricHeaders.map(() => null);
    for (const bucket of edctBuckets) {
      const edctValues = edctIndicators.map((ind) => (ind === bucket.indicator ? bucket.value : null));
      ws.addRow([
        bucket.state, '', bucket.facility,
        bucket.period, '',
        ...nullMetrics,
        ...edctValues,
      ]).commit();
    }

    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    ws.commit();
    this.logger.info(`  Monthly sheet: ${facilities.length} target facilities + ${edctBuckets.length} EDCT rows`);
  }

  // ── Helpers ───────────────────────────────────────────────────────────────────

  private findMaxDateRow(rows: TargetRow[], start: Dayjs, end: Dayjs): TargetRow | null {
    let selected: TargetRow | null = null;
    for (const row of rows) {
      if (
        (row.date.isSame(start, 'day') || row.date.isAfter(start)) &&
        (row.date.isSame(end,   'day') || row.date.isBefore(end))
      ) {
        if (!selected || row.date.isAfter(selected.date)) selected = row;
      }
    }
    return selected;
  }

  private coerceDashboardRow(values: string[], headers: string[]): unknown[] {
    return values.map((value, idx) => {
      const header = headers[idx];
      const text   = this.unwrapExcelText(value);
      if (['Value', 'Numerator', 'Denominator', 'Target'].includes(header)) {
        const n = Number(text.replace(/,/g, ''));
        return Number.isFinite(n) ? n : null;
      }
      if (header === 'AchievementPct') {
        const n = Number(text.replace('%', '').replace(/,/g, ''));
        return Number.isFinite(n) ? n / 100 : null;
      }
      return text;
    });
  }

  private parseCsvLine(line: string): string[] {
    const result: string[] = [];
    let current = '';
    let quoted  = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (quoted && line[i + 1] === '"') { current += '"'; i++; }
        else { quoted = !quoted; }
      } else if (ch === ',' && !quoted) {
        result.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
    result.push(current);
    return result;
  }

  private unwrapExcelText(value: string): string {
    const match = value.match(/^="(.*)"$/);
    return match ? match[1] : value;
  }

  private parseExcelDate(value: unknown): Dayjs | null {
    if (value instanceof Date) return dayjs(value).startOf('day');
    if (typeof value === 'number') return dayjs('1899-12-30').add(value, 'day').startOf('day');
    const parsed = dayjs(String(value ?? '').trim(), ['YYYY-MM-DD', 'DD/MM/YYYY', 'MM/DD/YYYY'], true);
    return parsed.isValid() ? parsed.startOf('day') : null;
  }

  private stripStatePrefix(value: string): string {
    return String(value ?? '').replace(/^(ad|bo|yo|ta)\s+/i, '').trim();
  }

  private normalizeFacility(value: string): string {
    return this.stripStatePrefix(value)
      .toLowerCase()
      .replace(/geidam/g, 'geidem')
      .replace(/rapha/g, 'rapah')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim()
      .replace(/\s+/g, ' ');
  }

  private widthForHeader(header: string): number {
    if (header === 'Facility' || header === 'FACILITY') return 32;
    if (header === 'Indicator')   return 24;
    if (header.length > 24)       return 24;
    return Math.max(12, Math.min(20, header.length + 2));
  }

  private styleHeaderRow(row: ExcelJS.Row): void {
    row.eachCell((cell) => {
      cell.font      = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border    = { bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } } };
    });
    row.height = 28;
  }
}
