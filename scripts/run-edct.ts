// Runs only the EDCT query (scripts/edct_query.sql) and writes outputs/EDCT_Only.csv.
// Uses the same POSTGRES_CONNECTION_STRING and CUSTOM_START/CUSTOM_END logic as index.ts.
//   npx ts-node scripts/run-edct.ts
import 'dotenv/config';
import path from 'path';
import fs from 'fs';
import winston from 'winston';
import { EdctService } from '../services/EdctService';
import { DateHelper, DateModeConfig } from '../helpers/DateHelper';

const logger = winston.createLogger({
  level: 'info',
  format: winston.format.printf(({ level, message }) => `[${level.toUpperCase()}] ${message}`),
  transports: [new winston.transports.Console()],
});

function csvCell(v: unknown): string {
  const s = v == null ? '' : String(v);
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

async function main(): Promise<void> {
  const connStr = process.env['POSTGRES_CONNECTION_STRING'];
  if (!connStr) throw new Error('POSTGRES_CONNECTION_STRING not set');

  const dateModeConfig: DateModeConfig = {
    mode: (process.env['DATE_MODE'] as DateModeConfig['mode']) ?? 'DAILY',
    referenceDate: process.env['REFERENCE_DATE'],
    customStart: process.env['CUSTOM_START'],
    customEnd: process.env['CUSTOM_END'],
  };

  const service = new EdctService(connStr, logger);
  if (!(await service.testConnection())) throw new Error('PostgreSQL connection failed');

  const sqlPath = path.resolve(process.cwd(), 'scripts', 'edct_query.sql');
  const range = new DateHelper(dateModeConfig).getGroupByDateRange();
  const rows = await service.fetchDashboardRows(sqlPath, range);

  const headers = ['Period', 'State', 'Facility', 'DATIMCode', 'Indicator', 'Disaggregation', 'Category', 'Sex', 'AgeBand', 'Value'] as const;
  const lines = [headers.join(',')];
  // Same ="..." text guard as OutputWriter.csvText, so Excel doesn't turn "1-4" into 4-Jan
  const textCols = new Set<string>(['Category', 'AgeBand']);
  for (const r of rows) {
    lines.push(headers.map(h => textCols.has(h) && r[h] != null ? csvCell(`="${r[h]}"`) : csvCell(r[h])).join(','));
  }

  const outDir = path.resolve(process.cwd(), 'outputs');
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, 'EDCT_Only.csv');
  fs.writeFileSync(outPath, '\uFEFF' + lines.join('\r\n') + '\r\n', 'utf8');

  const byIndicator = new Map<string, number>();
  for (const r of rows) byIndicator.set(r.Indicator, (byIndicator.get(r.Indicator) ?? 0) + (r.Value ?? 0));
  logger.info(`Wrote ${rows.length} rows to ${outPath}`);
  for (const [k, v] of byIndicator) logger.info(`  ${k}: ${v}`);
}

main().catch(err => { logger.error(String(err)); process.exit(1); });
