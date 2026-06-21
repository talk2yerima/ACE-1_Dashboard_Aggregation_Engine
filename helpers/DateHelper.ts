import dayjs, { Dayjs } from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';
import isBetween from 'dayjs/plugin/isBetween';
import isSameOrAfter from 'dayjs/plugin/isSameOrAfter';
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore';
import weekOfYear from 'dayjs/plugin/weekOfYear';
import quarterOfYear from 'dayjs/plugin/quarterOfYear';

dayjs.extend(customParseFormat);
dayjs.extend(isBetween);
dayjs.extend(isSameOrAfter);
dayjs.extend(isSameOrBefore);
dayjs.extend(weekOfYear);
dayjs.extend(quarterOfYear);

export type DateMode = 'TODAY' | 'DAILY' | 'WEEKLY' | 'MONTHLY' | 'QUARTERLY' | 'YEARLY' | 'CUSTOM';

export interface DateRange {
  start: Dayjs;
  end: Dayjs;
}

export interface DateModeConfig {
  mode: DateMode;
  referenceDate?: string;  // for CUSTOM mode: 'YYYY-MM-DD'
  customStart?: string;
  customEnd?: string;
}

const DATE_FORMATS = [
  'YYYY-MM-DD',
  'DD/MM/YYYY',
  'MM/DD/YYYY',
  'DD-MM-YYYY',
  'YYYY/MM/DD',
  'D/M/YYYY',
  'DD-MMM-YYYY',
];

export class DateHelper {
  private config: DateModeConfig;
  private today: Dayjs;

  constructor(config: DateModeConfig) {
    this.config = config;
    this.today = dayjs(config.referenceDate ?? undefined).startOf('day');
  }

  /** Parse a raw cell value into a Dayjs instance, trying multiple formats */
  parse(value: unknown): Dayjs | null {
    if (!value) return null;

    if (value instanceof Date) {
      const d = dayjs(value);
      return d.isValid() ? d : null;
    }

    if (typeof value === 'number') {
      // Excel serial date number
      const d = dayjs(new Date(Math.round((value - 25569) * 86400 * 1000)));
      return d.isValid() ? d : null;
    }

    const str = String(value).trim();

    // Excel serial date stored as a string (e.g. "45306") — happens when a
    // date cell value is converted via String() before being stored in rawComponents.
    // Valid Excel serials are roughly 1 (Jan 1 1900) to 2,958,465 (Dec 31 9999).
    const numericStr = parseFloat(str);
    if (!isNaN(numericStr) && numericStr >= 1 && numericStr <= 2958465 && String(numericStr) === str) {
      const d = dayjs(new Date(Math.round((numericStr - 25569) * 86400 * 1000)));
      if (d.isValid()) return d;
    }
    if (!str || str === '' || str.toLowerCase() === 'null') return null;

    for (const fmt of DATE_FORMATS) {
      const d = dayjs(str, fmt, true);
      if (d.isValid()) return d;
    }

    // Fallback: native parse
    const d = dayjs(str);
    return d.isValid() ? d : null;
  }

  /** Get the effective date range for the current mode */
  getRange(): DateRange {
    const { mode, customStart, customEnd } = this.config;
    const t = this.today;

    switch (mode) {
      case 'TODAY':
        return { start: t.startOf('day'), end: t.endOf('day') };

      case 'DAILY':
        return { start: t.startOf('day'), end: t.endOf('day') };

      case 'WEEKLY':
        return { start: t.startOf('week'), end: t.endOf('week') };

      case 'MONTHLY':
        return { start: t.startOf('month'), end: t.endOf('month') };

      case 'QUARTERLY':
        return { start: t.startOf('quarter'), end: t.endOf('quarter') };

      case 'YEARLY':
        return { start: t.startOf('year'), end: t.endOf('year') };

      case 'CUSTOM': {
        if (!customStart || !customEnd) {
          throw new Error('CUSTOM date mode requires customStart and customEnd');
        }
        const s = dayjs(customStart, 'YYYY-MM-DD', true);
        const e = dayjs(customEnd, 'YYYY-MM-DD', true);
        if (!s.isValid() || !e.isValid()) {
          throw new Error(`Invalid custom date range: ${customStart} - ${customEnd}`);
        }
        return { start: s.startOf('day'), end: e.endOf('day') };
      }

      default:
        throw new Error(`Unknown date mode: ${mode}`);
    }
  }

  /** Check whether a raw cell value falls within the current date range */
  isInRange(value: unknown): boolean {
    const d = this.parse(value);
    if (!d) return false;
    const range = this.getRange();
    return d.isBetween(range.start, range.end, 'day', '[]');
  }

  /** Return the period label formatted for display */
  getPeriodLabel(): string {
    const t = this.today;
    const { mode } = this.config;

    switch (mode) {
      case 'TODAY':
      case 'DAILY':
        return t.format('DD/MM/YYYY');
      case 'WEEKLY':
        return `W${String(t.week()).padStart(2, '0')}/${t.format('YYYY')}`;
      case 'MONTHLY':
        return t.format('MMM YYYY');
      case 'QUARTERLY':
        return `Q${t.quarter()} ${t.format('YYYY')}`;
      case 'YEARLY':
        return t.format('YYYY');
      case 'CUSTOM': {
        const range = this.getRange();
        return `${range.start.format('DD/MM/YYYY')} - ${range.end.format('DD/MM/YYYY')}`;
      }
      default:
        return t.format('DD/MM/YYYY');
    }
  }

  getToday(): Dayjs {
    return this.today;
  }

  /**
   * The date range used to filter rows in group-by-date modes (DAILY / WEEKLY /
   * MONTHLY / QUARTERLY).  In these modes dateMode filters are skipped and each
   * row gets its own period label, so we need a separate range gate.
   *
   * Priority:
   *   1. CUSTOM_START / CUSTOM_END env vars (explicit override)
   *   2. PEPFAR fiscal year start (Oct 1 of the current or previous year) → today
   */
  getGroupByDateRange(): DateRange {
    const { customStart, customEnd } = this.config;

    if (customStart) {
      const s = dayjs(customStart, 'YYYY-MM-DD', true);
      const e = customEnd
        ? dayjs(customEnd, 'YYYY-MM-DD', true).endOf('day')
        : this.today.endOf('day');
      if (s.isValid()) return { start: s.startOf('day'), end: e };
    }

    // Default: PEPFAR fiscal year start (Oct 1) → today
    return { start: this.getPEPFARFiscalYearStart(), end: this.today.endOf('day') };
  }

  /** Oct 1 of the current PEPFAR fiscal year. */
  getPEPFARFiscalYearStart(): Dayjs {
    // month() is 0-indexed: September = 8, October = 9
    const oct1ThisYear = this.today.startOf('year').add(9, 'month');
    return this.today.month() >= 9
      ? oct1ThisYear                         // Oct–Dec: FY started this year
      : oct1ThisYear.subtract(1, 'year');    // Jan–Sep: FY started last year
  }

  /** Full PEPFAR fiscal year: Oct 1 → Sep 30. */
  getPEPFARFiscalYearRange(): DateRange {
    const start = this.getPEPFARFiscalYearStart();
    const end = start.add(1, 'year').subtract(1, 'day').endOf('day');
    return { start, end };
  }

  /** PEPFAR FY first semi-annual period: Oct 1 → Mar 31. */
  getPreviousSemiQuarterRange(): DateRange {
    const start = this.getPEPFARFiscalYearStart();                        // Oct 1
    const end = start.add(6, 'month').subtract(1, 'day').endOf('day');   // Mar 31
    return { start, end };
  }

  /** PEPFAR FY second semi-annual period: Apr 1 → Sep 30. */
  getCurrentSemiQuarterRange(): DateRange {
    const start = this.getPEPFARFiscalYearStart().add(6, 'month');        // Apr 1
    const end = start.add(6, 'month').subtract(1, 'day').endOf('day');   // Sep 30
    return { start, end };
  }

  /** ART VL eligibility: any ARTStartDate up to (today - 180 days). */
  getBefore180DaysAgoRange(): DateRange {
    return {
      start: dayjs('1900-01-01').startOf('day'),
      end: this.today.subtract(180, 'day').endOf('day'),
    };
  }
}
