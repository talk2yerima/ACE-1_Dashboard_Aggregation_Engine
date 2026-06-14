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
    | 'greaterThan'
    | 'lessThan'
    | 'greaterThanOrEqual'
    | 'lessThanOrEqual'
    | 'isNotNull'
    | 'isNull';
  value?: unknown;
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
