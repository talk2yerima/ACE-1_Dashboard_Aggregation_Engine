import path from 'path';
import fs from 'fs';
import winston from 'winston';

import { AggregationEngine } from './services/AggregationEngine';
import { OutputWriter } from './services/OutputWriter';
import { DateModeConfig } from './helpers/DateHelper';

// ─── Logger Setup ────────────────────────────────────────────────────────────

const outputDir = path.resolve(process.cwd(), 'outputs');
fs.mkdirSync(outputDir, { recursive: true });

const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
    winston.format.printf(({ timestamp, level, message }) => `[${timestamp}] [${level.toUpperCase()}] ${message}`),
  ),
  transports: [
    new winston.transports.Console(),
    new winston.transports.File({
      filename: path.join(outputDir, 'process.log'),
      options: { flags: 'w' },
    }),
  ],
});

// ─── Configuration ───────────────────────────────────────────────────────────

// Date mode can be overridden via environment variables:
//   DATE_MODE=MONTHLY node dist/index.js
//   DATE_MODE=CUSTOM CUSTOM_START=2026-01-01 CUSTOM_END=2026-03-31 node dist/index.js

const dateModeConfig: DateModeConfig = {
  mode: (process.env['DATE_MODE'] as DateModeConfig['mode']) ?? 'DAILY',
  referenceDate: process.env['REFERENCE_DATE'],   // e.g. '2026-06-13'
  customStart: process.env['CUSTOM_START'],
  customEnd: process.env['CUSTOM_END'],
};

const inputDir = path.resolve(process.cwd(), 'input');
const workbookPath = process.env['RADET_FILE']
  ? path.resolve(process.env['RADET_FILE'])
  : (() => {
      const match = fs.readdirSync(inputDir).find(f => f.startsWith('ACE-1_Combined_RADET') && f.endsWith('.xlsx'));
      return match ? path.join(inputDir, match) : path.join(inputDir, 'RADET.xlsx');
    })();

const configDir = path.resolve(process.cwd(), 'config');

// ─── Main ────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  logger.info('═══════════════════════════════════════════════════');
  logger.info('  RADET Dashboard Aggregation Engine  v1.0.0');
  logger.info('═══════════════════════════════════════════════════');
  logger.info(`Date mode:      ${dateModeConfig.mode}`);
  logger.info(`Workbook:       ${workbookPath}`);
  logger.info(`Config dir:     ${configDir}`);
  logger.info(`Output dir:     ${outputDir}`);

  if (!fs.existsSync(workbookPath)) {
    logger.error(`Workbook not found: ${workbookPath}`);
    logger.error('Place an ACE-1_Combined_RADET*.xlsx file in the input/ folder, or set RADET_FILE env var.');
    process.exit(1);
  }

  // ── Aggregation ────────────────────────────────────────────────────────
  const engine = new AggregationEngine({
    workbookPath,
    configDir,
    dateModeConfig,
    logger,
  });

  await engine.init();
  const dashboardRows = await engine.process();

  // ── Output ─────────────────────────────────────────────────────────────
  const writer = new OutputWriter({ outputDir, logger });

  if (dashboardRows.length > 0) {
    const { xlsx, csv } = await writer.writeDashboard(dashboardRows);
    logger.info(`\nOutput files:`);
    logger.info(`  ${xlsx}`);
    logger.info(`  ${csv}`);
  } else {
    logger.warn('No dashboard rows generated — check your workbook and date mode.');
  }

  // ── Validation Report ──────────────────────────────────────────────────
  if (engine.validationIssues.length > 0) {
    const reportPath = await writer.writeValidationReport(engine.validationIssues);
    logger.warn(`\nValidation issues found. Report: ${reportPath}`);
  }

  // ── Final Stats ────────────────────────────────────────────────────────
  logger.info('\n─── Processing Summary ───────────────────────────');
  logger.info(`Rows read:             ${engine.stats.totalRows}`);
  logger.info(`Rows after filtering:  ${engine.stats.filteredRows}`);
  logger.info(`Dashboard rows out:    ${engine.stats.aggregatedRows}`);
  logger.info(`Duration:              ${engine.stats.durationMs}ms`);
  logger.info(`Indicators done:       ${engine.stats.indicatorsProcessed.join(', ')}`);
  logger.info('──────────────────────────────────────────────────\n');
}

main().catch((err) => {
  logger.error(`Fatal error: ${String(err)}`);
  process.exit(1);
});
