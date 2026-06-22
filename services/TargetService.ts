import ExcelJS from 'exceljs';
import dayjs, { Dayjs } from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore';
import fs from 'fs';
import path from 'path';
import yaml from 'js-yaml';
import { Logger } from 'winston';
import { DashboardRow } from './AggregationEngine';

dayjs.extend(customParseFormat);
dayjs.extend(isSameOrBefore);

interface OrgUnitsConfig {
  orgUnits?: {
    datimCodes?: Record<string, { facility?: string; state?: string }>;
  };
}

type TargetPart = { column: string; ageGroup?: '<15' | '15+'; sex?: 'Female' | 'Male' };

interface TargetEntry {
  date: Dayjs;
  value: number;
}

export interface TargetLookupResult {
  target: number;
  achievementPct: number | null;
  targetDate: string;
}

export class TargetService {
  private logger: Logger;
  private targetByKey = new Map<string, TargetEntry[]>();
  private targetColumns = new Map<string, number>();
  private facilityToDatim = new Map<string, string>();

  constructor(logger: Logger) {
    this.logger = logger;
  }

  async load(targetPath: string, configDir: string): Promise<void> {
    if (!fs.existsSync(targetPath)) {
      this.logger.warn(`TargetService: target workbook not found: ${targetPath}`);
      return;
    }

    this.loadFacilityMap(configDir);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(targetPath);
    const sheet = workbook.worksheets[0];
    if (!sheet) {
      this.logger.warn(`TargetService: no worksheets found in ${targetPath}`);
      return;
    }

    const headerRow = sheet.getRow(1);
    headerRow.eachCell({ includeEmpty: false }, (cell, col) => {
      const header = String(cell.value ?? '').trim();
      if (header) this.targetColumns.set(header, col);
    });

    const facilityCol = this.targetColumns.get('FACILITY') ?? 3;
    const dateCol = this.targetColumns.get('DATE') ?? 5;
    let rowsIndexed = 0;

    for (let rowNum = 2; rowNum <= sheet.rowCount; rowNum++) {
      const row = sheet.getRow(rowNum);
      const targetDate = this.parseDate(row.getCell(dateCol).value);
      if (!targetDate) continue;

      const facilityRaw = String(row.getCell(facilityCol).value ?? '').trim();
      const datim = this.facilityToDatim.get(this.normalizeFacility(facilityRaw));
      if (!datim) continue;

      for (const [indicator, parts] of Object.entries(TARGET_MAP)) {
        for (const part of parts) {
          const col = this.targetColumns.get(part.column);
          if (!col) continue;
          const value = this.toNumber(row.getCell(col).value);
          if (value === null) continue;
          const key = this.makeKey(datim, indicator, part.ageGroup ?? '', part.sex ?? '');
          this.addTarget(key, targetDate, value);
        }
      }
      rowsIndexed++;
    }

    for (const entries of this.targetByKey.values()) {
      entries.sort((a, b) => a.date.valueOf() - b.date.valueOf());
    }

    this.logger.info(
      `TargetService: loaded ${rowsIndexed.toLocaleString()} target rows from ` +
      `${path.basename(targetPath)} (${this.targetByKey.size.toLocaleString()} lookup keys)`,
    );
  }

  apply(row: DashboardRow): DashboardRow {
    const result = this.getTarget(row);
    if (!result) return row;

    return {
      ...row,
      Target: result.target,
      AchievementPct: result.achievementPct,
    };
  }

  getTarget(row: DashboardRow): TargetLookupResult | null {
    const period = this.parsePeriod(row.Period);
    if (!period || !row.DATIMCode || !row.Indicator) return null;

    const ageGroup = this.toTargetAgeGroup(row.AgeBand);
    const sex = this.toTargetSex(row.Sex);
    const indicatorParts = TARGET_MAP[row.Indicator];
    if (!indicatorParts) return null;

    const ageSexParts = this.resolveParts(row.Indicator, ageGroup, sex, true);
    const totalParts = this.resolveParts(row.Indicator, ageGroup, sex, false);
    const ageSexResult = this.sumTargets(row.DATIMCode, row.Indicator, period, ageSexParts);
    const totalResult = ageSexResult.matched
      ? ageSexResult
      : this.sumTargets(row.DATIMCode, row.Indicator, period, totalParts);

    if (!totalResult.matched || !totalResult.targetDate) return null;
    const target = totalResult.target;

    const achievementPct =
      target > 0 && row.Value !== null && row.Value !== undefined
        ? parseFloat((((row.Value ?? 0) / target) * 100).toFixed(2))
        : null;

    return { target, achievementPct, targetDate: totalResult.targetDate };
  }

  private resolveParts(
    indicator: string,
    ageGroup: '<15' | '15+' | '',
    sex: 'Female' | 'Male' | '',
    ageSexOnly: boolean,
  ): TargetPart[] {
    const parts = TARGET_MAP[indicator] ?? [];
    const ageSexParts = parts.filter(p => p.ageGroup || p.sex);
    if (!ageSexOnly) return parts.filter(p => !p.ageGroup && !p.sex);
    if (!ageGroup || !sex) return [];
    return ageSexParts.filter(p => p.ageGroup === ageGroup && p.sex === sex);
  }

  private sumTargets(
    datim: string,
    indicator: string,
    period: Dayjs,
    parts: TargetPart[],
  ): { matched: boolean; target: number; targetDate: string | null } {
    let target = 0;
    let matched = false;
    let targetDate: string | null = null;
    for (const part of parts) {
      const key = this.makeKey(datim, indicator, part.ageGroup ?? '', part.sex ?? '');
      const entry = this.findTargetOnOrBefore(key, period);
      if (entry !== null) {
        target += entry.value;
        targetDate = entry.date.format('YYYY-MM-DD');
        matched = true;
      }
    }
    return { matched, target, targetDate };
  }

  private addTarget(key: string, date: Dayjs, value: number): void {
    const entries = this.targetByKey.get(key) ?? [];
    const existing = entries.find(e => e.date.isSame(date, 'day'));
    if (existing) {
      existing.value += value;
    } else {
      entries.push({ date, value });
      this.targetByKey.set(key, entries);
    }
  }

  private findTargetOnOrBefore(key: string, period: Dayjs): TargetEntry | null {
    const entries = this.targetByKey.get(key);
    if (!entries || entries.length === 0) return null;

    let lo = 0;
    let hi = entries.length - 1;
    let found: TargetEntry | null = null;
    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2);
      if (entries[mid].date.isSameOrBefore(period, 'day')) {
        found = entries[mid];
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return found;
  }

  private loadFacilityMap(configDir: string): void {
    const orgPath = path.join(configDir, 'orgUnits.yaml');
    const config = yaml.load(fs.readFileSync(orgPath, 'utf8')) as OrgUnitsConfig;
    const datimCodes = config.orgUnits?.datimCodes ?? {};

    for (const [datim, info] of Object.entries(datimCodes)) {
      if (info.facility) this.facilityToDatim.set(this.normalizeFacility(info.facility), datim);
    }

    this.facilityToDatim.set(this.normalizeFacility('yo Geidam General Hospital'), 'JcYmXdOSndf');
    this.facilityToDatim.set(this.normalizeFacility('ta Rapha Hospital'), 'rmGHrOrVW9r');
  }

  private normalizeFacility(value: string): string {
    return String(value ?? '')
      .toLowerCase()
      .replace(/^(ad|bo|yo|ta)\s+/, '')
      .replace(/geidam/g, 'geidem')
      .replace(/rapha/g, 'rapah')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim()
      .replace(/\s+/g, ' ');
  }

  private toTargetAgeGroup(ageBand: string): '<15' | '15+' | '' {
    if (!ageBand) return '';
    if (['<1', '1-4', '5-9', '10-14'].includes(ageBand)) return '<15';
    if (['15-19', '20-24', '25-29', '30-34', '35-39', '40-44', '45-49', '50+'].includes(ageBand)) return '15+';
    return '';
  }

  private toTargetSex(sex: string): 'Female' | 'Male' | '' {
    const normalized = String(sex ?? '').trim().toLowerCase();
    if (normalized === 'female') return 'Female';
    if (normalized === 'male') return 'Male';
    return '';
  }

  private parsePeriod(value: unknown): Dayjs | null {
    const parsed = dayjs(String(value ?? '').trim(), ['DD/MM/YYYY', 'YYYY-MM-DD'], true);
    return parsed.isValid() ? parsed.startOf('day') : null;
  }

  private parseDate(value: unknown): Dayjs | null {
    if (value instanceof Date) return dayjs(value).startOf('day');
    if (typeof value === 'number') {
      const excelEpoch = dayjs('1899-12-30');
      return excelEpoch.add(value, 'day').startOf('day');
    }
    const parsed = dayjs(String(value ?? '').trim(), ['YYYY-MM-DD', 'DD/MM/YYYY', 'MM/DD/YYYY'], true);
    return parsed.isValid() ? parsed.startOf('day') : null;
  }

  private toNumber(value: unknown): number | null {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (value && typeof value === 'object' && 'result' in value) {
      return this.toNumber((value as ExcelJS.CellFormulaValue).result);
    }
    const parsed = parseFloat(String(value ?? '').replace(/,/g, '').trim());
    return Number.isFinite(parsed) ? parsed : null;
  }

  private makeKey(datim: string, indicator: string, ageGroup: string, sex: string): string {
    return `${datim}|${indicator}|${ageGroup}|${sex}`;
  }
}

const TARGET_MAP: Record<string, TargetPart[]> = {
  HTS_TST: [
    { column: 'HTS_TST.Neg.T' },
    { column: 'HTS_TST.Pos.T' },
    { column: 'HTS_TST_NEG<15Female', ageGroup: '<15', sex: 'Female' },
    { column: 'HTS_TST_NEG<15Male', ageGroup: '<15', sex: 'Male' },
    { column: 'HTS_TST_NEG15+Female', ageGroup: '15+', sex: 'Female' },
    { column: 'HTS_TST_NEG15+Male', ageGroup: '15+', sex: 'Male' },
    { column: 'HTS_TST_POS<15Female', ageGroup: '<15', sex: 'Female' },
    { column: 'HTS_TST_POS<15Male', ageGroup: '<15', sex: 'Male' },
    { column: 'HTS_TST_POS15+Female', ageGroup: '15+', sex: 'Female' },
    { column: 'HTS_TST_POS15+Male', ageGroup: '15+', sex: 'Male' },
  ],
  HTS_TST_POS: [
    { column: 'HTS_TST.Pos.T' },
    { column: 'HTS_TST_POS<15Female', ageGroup: '<15', sex: 'Female' },
    { column: 'HTS_TST_POS<15Male', ageGroup: '<15', sex: 'Male' },
    { column: 'HTS_TST_POS15+Female', ageGroup: '15+', sex: 'Female' },
    { column: 'HTS_TST_POS15+Male', ageGroup: '15+', sex: 'Male' },
  ],
  TX_New: targetAgeSex('TX_NEW', 'TX_NEW.T'),
  TX_CURR: targetAgeSex('TX_CURR', 'TX_CURR.T'),
  'VL Result Received_PVLSD': targetAgeSex('TX_PVLS_D', 'TX_PVLS.D.T'),
  PVLSN: targetAgeSex('TX_PVLS_N', 'TX_PVLS.N.T'),
  PMTCT_ART_Already: [
    { column: 'PMTCT_ART.Already.T' },
    { column: 'PMTCT_ART_Already15+Female', ageGroup: '15+', sex: 'Female' },
  ],
  PMTCT_ART_New: [
    { column: 'PMTCT_ART.New.T' },
    { column: 'PMTCT_ART_New<15Female', ageGroup: '<15', sex: 'Female' },
    { column: 'PMTCT_ART_New15+Female', ageGroup: '15+', sex: 'Female' },
  ],
  PMTCT_STAT_New_Neg: [
    { column: 'PMTCT_STAT.N.New.Neg.T' },
    { column: 'PMTCT_STAT_New_Negative<15Female', ageGroup: '<15', sex: 'Female' },
    { column: 'PMTCT_STAT_New_Negative15+Female', ageGroup: '15+', sex: 'Female' },
  ],
  PMTCT_STAT_New_Pos: [
    { column: 'PMTCT_STAT.N.New.Pos.T' },
    { column: 'PMTCT_STAT_New_Positive<15Female', ageGroup: '<15', sex: 'Female' },
    { column: 'PMTCT_STAT_New_Positive15+Female', ageGroup: '15+', sex: 'Female' },
  ],
  TB_PREV_N_Already: targetAgeSex('TB_PREV_Already', 'TB_PREV.N.Already.T'),
  TB_PREV_N_New: targetAgeSex('TB_PREV_New', 'TB_PREV.N.New.T'),
  TX_TB_N_Already: targetAgeSex('TX_TB_Already', 'TX_TB.N.Already.T'),
  TX_TB_N_New: targetAgeSex('TX_TB_New', 'TX_TB.N.New.T'),
};

function targetAgeSex(prefix: string, totalColumn: string): TargetPart[] {
  return [
    { column: totalColumn },
    { column: `${prefix}<15Female`, ageGroup: '<15', sex: 'Female' },
    { column: `${prefix}<15Male`, ageGroup: '<15', sex: 'Male' },
    { column: `${prefix}15+Female`, ageGroup: '15+', sex: 'Female' },
    { column: `${prefix}15+Male`, ageGroup: '15+', sex: 'Male' },
  ];
}
