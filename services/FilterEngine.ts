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
    | 'isNull';
  value?: unknown;
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
    return parseFloat(String(val));
  }
}
