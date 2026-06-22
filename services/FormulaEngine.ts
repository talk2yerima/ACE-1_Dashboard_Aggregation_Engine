import { Logger } from 'winston';
import { DashboardRow } from './AggregationEngine';

export interface FormulaIndicatorDef {
  name: string;
  description?: string;
  numerator: string;
  denominator: string;
  formula: 'numerator / denominator' | 'numerator - denominator' | 'numerator + denominator' | 'numerator * denominator';
  outputType: 'percentage' | 'count' | 'ratio';
  groupBy: string[];
}

export interface FormulaEngineOptions {
  logger: Logger;
}

export class FormulaEngine {
  private logger: Logger;

  constructor(options: FormulaEngineOptions) {
    this.logger = options.logger;
  }

  /**
   * Calculate formula indicators from existing dashboard rows.
   */
  calculate(
    formulaDefs: FormulaIndicatorDef[],
    existingRows: DashboardRow[],
    period: string,
  ): DashboardRow[] {
    const results: DashboardRow[] = [];

    for (const def of formulaDefs) {
      this.logger.info(`FormulaEngine: calculating ${def.name}`);

      const numRows = existingRows.filter((r) => r.Indicator === def.numerator);
      const denRows = existingRows.filter((r) => r.Indicator === def.denominator);

      if (numRows.length === 0) {
        this.logger.warn(`FormulaEngine: no rows for numerator indicator ${def.numerator}`);
        continue;
      }

      // Index denominator rows by group key for fast lookup
      const denIndex = new Map<string, number>();
      for (const row of denRows) {
        const key = this.makeKey(row, def.groupBy);
        denIndex.set(key, (denIndex.get(key) ?? 0) + (row.Value ?? 0));
      }

      // Group numerator rows
      const numIndex = new Map<string, { value: number; row: DashboardRow }>();
      for (const row of numRows) {
        const key = this.makeKey(row, def.groupBy);
        const existing = numIndex.get(key);
        if (existing) {
          existing.value += row.Value ?? 0;
        } else {
          numIndex.set(key, { value: row.Value ?? 0, row });
        }
      }

      for (const [key, { value: numVal, row: templateRow }] of numIndex.entries()) {
        const denVal = denIndex.get(key) ?? 0;
        const computed = this.compute(def.formula, numVal, denVal);

        const outputRow: DashboardRow = {
          Period: period,
          State: templateRow.State,
          Facility: templateRow.Facility,
          DATIMCode: templateRow.DATIMCode,
          Indicator: def.name,
          Disaggregation: def.outputType === 'percentage' ? 'Rate' : 'Count',
          Category: def.name,
          Sex: '',
          AgeBand: '',
          Value: def.outputType === 'percentage' ? parseFloat((computed * 100).toFixed(2)) : computed,
          Numerator: numVal,
          Denominator: denVal,
          Target: null,
          AchievementPct: null,
        };

        results.push(outputRow);
      }

      this.logger.info(`FormulaEngine: ${def.name} → ${results.length} rows`);
    }

    return results;
  }

  private compute(
    formula: FormulaIndicatorDef['formula'],
    numerator: number,
    denominator: number,
  ): number {
    switch (formula) {
      case 'numerator / denominator':
        return denominator === 0 ? 0 : numerator / denominator;
      case 'numerator - denominator':
        return numerator - denominator;
      case 'numerator + denominator':
        return numerator + denominator;
      case 'numerator * denominator':
        return numerator * denominator;
      default:
        return 0;
    }
  }

  private makeKey(row: DashboardRow, groupBy: string[]): string {
    const r = row as unknown as Record<string, unknown>;
    return groupBy.map((k) => String(r[k] ?? '')).join('|||');
  }
}
