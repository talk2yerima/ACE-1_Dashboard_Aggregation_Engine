import { DateHelper, DateModeConfig } from '../helpers/DateHelper';
import { Logger } from 'winston';

export interface FilterDef {
  column: string;
  operator:
    | 'equals'
    | 'notEquals'
    | 'contains'
    | 'notContains'
    | 'inList'
    | 'notInList'
    | 'dateMode'
    | 'dateinRange'
    | 'dateNotInRange'
    | 'dateDiff'
    | 'greaterThan'
    | 'lessThan'
    | 'greaterThanOrEqual'
    | 'lessThanOrEqual'
    | 'isNotNull'
    | 'isNull'
    | 'equalsColumn'
    | 'notEqualsColumn'
    | 'viralLoadSuppressed';
  value?: unknown;
  /** For equalsColumn / notEqualsColumn: the other column name to compare against */
  valueColumn?: string;
  ref?: string;
  caseSensitive?: boolean;
}

export interface FilterEngineOptions {
  dateModeConfig: DateModeConfig;
  logger: Logger;
}

export class FilterEngine {
  private dateHelpers: Map<string, DateHelper> = new Map();
  private globalDateConfig: DateModeConfig;
  private logger: Logger;

  constructor(options: FilterEngineOptions) {
    this.globalDateConfig = options.dateModeConfig;
    this.logger = options.logger;
  }

  private getDateHelper(column: string): DateHelper {
    if (!this.dateHelpers.has(column)) {
      this.dateHelpers.set(column, new DateHelper(this.globalDateConfig));
    }
    return this.dateHelpers.get(column)!;
  }

  /** Returns true if the row passes ALL filters */
  passesAll(row: Record<string, unknown>, filters: FilterDef[]): boolean {
    for (const filter of filters) {
      if (!this.passes(row, filter)) return false;
    }
    return true;
  }

  /** Returns true if the row passes AT LEAST ONE filter */
  passesAny(row: Record<string, unknown>, filters: FilterDef[]): boolean {
    for (const filter of filters) {
      if (this.passes(row, filter)) return true;
    }
    return false;
  }

  /** Returns true if the row passes a single filter */
  passes(row: Record<string, unknown>, filter: FilterDef): boolean {
    const raw = row[filter.column];
    const { operator, value, caseSensitive = false } = filter;

    switch (operator) {
      case 'isNull':
        return raw === null || raw === undefined || raw === '';

      case 'isNotNull':
        return raw !== null && raw !== undefined && raw !== '';

      case 'equals': {
        const a = this.normalizeStr(raw, caseSensitive);
        const b = this.normalizeStr(value, caseSensitive);
        return a === b;
      }

      case 'notEquals': {
        const a = this.normalizeStr(raw, caseSensitive);
        const b = this.normalizeStr(value, caseSensitive);
        return a !== b;
      }

      case 'contains': {
        const a = this.normalizeStr(raw, caseSensitive);
        const b = this.normalizeStr(value, caseSensitive);
        return a.includes(b);
      }

      case 'notContains': {
        const a = this.normalizeStr(raw, caseSensitive);
        const b = this.normalizeStr(value, caseSensitive);
        return !a.includes(b);
      }

      case 'inList': {
        if (!Array.isArray(value)) return false;
        const a = this.normalizeStr(raw, caseSensitive);
        return (value as string[]).some((v) => this.normalizeStr(v, caseSensitive) === a);
      }

      case 'notInList': {
        if (!Array.isArray(value)) return true;
        const a = this.normalizeStr(raw, caseSensitive);
        return !(value as string[]).some((v) => this.normalizeStr(v, caseSensitive) === a);
      }

      case 'dateMode': {
        const helper = this.getDateHelper(filter.column);
        return helper.isInRange(raw);
      }

      case 'dateinRange': {
        const helper = this.getDateHelper(filter.column);
        const d = helper.parse(raw);
        if (!d) return false;
        const range = value === 'CurrentFY'
          ? helper.getPEPFARFiscalYearRange()
          : value === 'PreviousSemiQuarter'
          ? helper.getPreviousSemiQuarterRange()
          : value === 'CurrentSemiQuarter'
          ? helper.getCurrentSemiQuarterRange()
          : value === 'Before180DaysAgo'
          ? helper.getBefore180DaysAgoRange()
          : helper.getRange();
        return d.isBetween(range.start, range.end, 'day', '[]');
      }

      case 'dateNotInRange': {
        const helper = this.getDateHelper(filter.column);
        const d = helper.parse(raw);
        if (!d) return false;
        const range = value === 'CurrentFY'
          ? helper.getPEPFARFiscalYearRange()
          : helper.getRange();
        return !d.isBetween(range.start, range.end, 'day', '[]');
      }

      case 'dateDiff': {
        const helper = this.getDateHelper(filter.column);
        const d1 = helper.parse(raw);
        const d2 = filter.ref ? helper.parse(row[filter.ref]) : null;
        if (!d1 || !d2) return false;
        const diffDays = Math.abs(d1.diff(d2, 'day'));
        const match = String(filter.value ?? '').trim().match(/^([<>]=?)\s*(\d+)$/);
        if (!match) return false;
        const [, op, numStr] = match;
        const n = parseInt(numStr, 10);
        if (op === '<')  return diffDays < n;
        if (op === '>')  return diffDays > n;
        if (op === '<=') return diffDays <= n;
        if (op === '>=') return diffDays >= n;
        return false;
      }

      case 'greaterThan':
        return this.toNum(raw) > this.toNum(value);

      case 'lessThan':
        return this.toNum(raw) < this.toNum(value);

      case 'greaterThanOrEqual':
        return this.toNum(raw) >= this.toNum(value);

      case 'lessThanOrEqual':
        return this.toNum(raw) <= this.toNum(value);

      case 'equalsColumn': {
        const other = filter.valueColumn ? row[filter.valueColumn] : undefined;
        return this.normalizeStr(raw, caseSensitive) === this.normalizeStr(other, caseSensitive);
      }

      case 'notEqualsColumn': {
        const other = filter.valueColumn ? row[filter.valueColumn] : undefined;
        return this.normalizeStr(raw, caseSensitive) !== this.normalizeStr(other, caseSensitive);
      }

      case 'viralLoadSuppressed':
        return this.isViralLoadSuppressed(raw);

      default:
        this.logger.warn(`Unknown filter operator: ${operator}`);
        return true;
    }
  }

  private normalizeStr(val: unknown, caseSensitive: boolean): string {
    const s = val === null || val === undefined ? '' : String(val).trim();
    return caseSensitive ? s : s.toLowerCase();
  }

  private toNum(val: unknown): number {
    if (typeof val === 'number') return val;
    const s = String(val ?? '').trim().toLowerCase().replace(/\s+/g, ' ');

    // All known text forms of "undetectable / below limit of detection"
    // map to 0 so numeric comparisons (lessThan 1000) work correctly.
    if (VL_SUPPRESSED_TEXT.has(s)) return 0;

    // "< N" or "<N" patterns (e.g. "<20", "< 30") — the VL is below N,
    // so treat as 0 for comparison purposes.
    const ltMatch = s.match(/^<\s*(\d+(?:\.\d+)?)$/);
    if (ltMatch) return 0;

    // Strip thousand separators (e.g. "1,936" → 1936) before parsing
    return parseFloat(s.replace(/,/g, ''));
  }

  /**
   * Returns true when a raw viral load value represents a suppressed or
   * undetectable result.  Handles all DHIS2 / lab system variants seen in
   * RADET exports:
   *   numeric:   0, 0.0, 0.00  (and any value < 1 000 copies/mL)
   *   less-than: <20, < 20, <30, < 30  (any "<N" where N ≤ 1 000)
   *   text:      Not Detected, NotDetected, TargetNotDetected,
   *              Target Not Detected, < Titermin, LDL, TND, Undetectable …
   */
  private isViralLoadSuppressed(val: unknown): boolean {
    if (val === null || val === undefined || val === '') return false;
    const s = String(val).trim().toLowerCase().replace(/\s+/g, ' ');

    if (VL_SUPPRESSED_TEXT.has(s)) return true;

    // "<N" patterns — suppressed if threshold ≤ 1 000
    const ltMatch = s.match(/^<\s*(\d+(?:\.\d+)?)$/);
    if (ltMatch) return parseFloat(ltMatch[1]) <= 1000;

    // Numeric — suppressed if < 1 000
    const n = parseFloat(s.replace(/,/g, ''));
    return !isNaN(n) && n < 1000;
  }
}

// ── Viral load: all known text representations of "undetectable / suppressed" ──
const VL_SUPPRESSED_TEXT = new Set([
  // Standard English
  'not detected',
  'notdetected',
  'target not detected',
  'targetnotdetected',
  'tnd',
  'ldl',                       // lower than detection level
  'undetectable',
  'undetected',
  'below detection',
  'bdl',                       // below detection limit
  'below limit of detection',
  'blod',
  // Lab-system variants
  '< titermin',                // below titration minimum
  '<titermin',
  'titermin',
  'below lod',
  'b.d.',
  'b.d',
  'nd',                        // not detected (abbreviated)
]);

