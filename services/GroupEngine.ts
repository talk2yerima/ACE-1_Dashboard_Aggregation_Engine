import _ from 'lodash';
import { Logger } from 'winston';

export interface GroupEngineOptions {
  logger: Logger;
}

export type AggregationMethod = 'COUNT' | 'SUM' | 'AVG' | 'MIN' | 'MAX' | 'COUNTDISTINCT';

export interface GroupResult {
  keys: Record<string, string>;
  count: number;
  sum?: number;
  avg?: number;
  min?: number;
  max?: number;
  distinctCount?: number;
}

export class GroupEngine {
  private logger: Logger;

  constructor(options: GroupEngineOptions) {
    this.logger = options.logger;
  }

  /**
   * Group rows by key columns and apply an aggregation method.
   * @param rows         Filtered rows to aggregate
   * @param groupByKeys  Column names to group by
   * @param method       Aggregation method
   * @param valueColumn  Column to aggregate (for SUM/AVG/MIN/MAX)
   */
  aggregate(
    rows: Record<string, unknown>[],
    groupByKeys: string[],
    method: AggregationMethod = 'COUNT',
    valueColumn?: string,
  ): GroupResult[] {
    if (rows.length === 0) return [];

    const grouped = _.groupBy(rows, (row) =>
      groupByKeys.map((k) => String(row[k] ?? '')).join('|||'),
    );

    const results: GroupResult[] = [];

    for (const [, group] of Object.entries(grouped)) {
      const keyRecord: Record<string, string> = {};
      for (const k of groupByKeys) {
        keyRecord[k] = String(group[0][k] ?? '');
      }

      const result: GroupResult = { keys: keyRecord, count: group.length };

      if (valueColumn) {
        const nums = group
          .map((r) => {
            const v = r[valueColumn];
            return typeof v === 'number' ? v : parseFloat(String(v));
          })
          .filter((n) => !isNaN(n));

        result.sum = nums.reduce((a, b) => a + b, 0);
        result.avg = nums.length > 0 ? result.sum / nums.length : 0;
        result.min = nums.length > 0 ? Math.min(...nums) : 0;
        result.max = nums.length > 0 ? Math.max(...nums) : 0;
        result.distinctCount = new Set(group.map((r) => String(r[valueColumn] ?? ''))).size;
      }

      results.push(result);
    }

    this.logger.debug(
      `GroupEngine: ${rows.length} rows → ${results.length} groups (keys: ${groupByKeys.join(', ')})`,
    );

    return results;
  }

  /** Get the value for the requested aggregation method */
  getValue(result: GroupResult, method: AggregationMethod): number {
    switch (method) {
      case 'COUNT':
        return result.count;
      case 'SUM':
        return result.sum ?? 0;
      case 'AVG':
        return result.avg ?? 0;
      case 'MIN':
        return result.min ?? 0;
      case 'MAX':
        return result.max ?? 0;
      case 'COUNTDISTINCT':
        return result.distinctCount ?? 0;
      default:
        return result.count;
    }
  }
}
