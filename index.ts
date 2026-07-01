import 'dotenv/config';
import path from 'path';
import fs from 'fs';
import winston from 'winston';
import { BlobServiceClient } from '@azure/storage-blob';

import { AggregationEngine, DashboardRow } from './services/AggregationEngine';
import { OutputWriter } from './services/OutputWriter';
import { TargetService } from './services/TargetService';
import { WorkbookReportWriter } from './services/WorkbookReportWriter';
import { EdctService } from './services/EdctService';
import { DateHelper, DateModeConfig } from './helpers/DateHelper';

// â”€â”€â”€ Logger Setup â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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

// â”€â”€â”€ Configuration â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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
fs.mkdirSync(inputDir, { recursive: true });

// â”€â”€ Resolve workbook path â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Priority:
//   1. RADET_FILE env var (explicit override)
//   2. Download most-recently-modified .xlsx from AZURE_RADET_SOURCE_CONTAINER
//   3. Local input/ folder fallback

// Persists ETag of the last downloaded blob so we only re-download when the file changes.
const etagCachePath = path.join(inputDir, '.blob-etag-cache.json');

function loadEtagCache(): Record<string, string> {
  try { return JSON.parse(fs.readFileSync(etagCachePath, 'utf8')); } catch { return {}; }
}

function saveEtagCache(cache: Record<string, string>): void {
  fs.writeFileSync(etagCachePath, JSON.stringify(cache, null, 2), 'utf8');
}

function hasZipEndMarker(filePath: string): boolean {
  if (!fs.existsSync(filePath)) return false;
  const stat = fs.statSync(filePath);
  if (stat.size < 22) return false;

  const fd = fs.openSync(filePath, 'r');
  try {
    const maxTail = Math.min(stat.size, 65_558);
    const tail = Buffer.alloc(maxTail);
    fs.readSync(fd, tail, 0, maxTail, stat.size - maxTail);
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail[i] === 0x50 && tail[i + 1] === 0x4b && tail[i + 2] === 0x05 && tail[i + 3] === 0x06) {
        return true;
      }
    }
    return false;
  } finally {
    fs.closeSync(fd);
  }
}

async function resolveWorkbookPath(): Promise<string> {
  if (process.env['RADET_FILE']) {
    return path.resolve(process.env['RADET_FILE']);
  }

  const connectionString = process.env['AZURE_STORAGE_CONNECTION_STRING'];
  const sourceContainer  = process.env['AZURE_RADET_SOURCE_CONTAINER'] ?? 'combine-radets-xls';

  if (connectionString) {
    try {
      logger.info(`Checking latest RADET file from blob storage (${sourceContainer})...`);
      const blobService = BlobServiceClient.fromConnectionString(connectionString);
      const container   = blobService.getContainerClient(sourceContainer);

      // List all .xlsx blobs; prefer ACE-1_Combined_RADET* files, then fall back to most recently modified
      type BlobEntry = { name: string; lastModified: Date; etag: string };
      let preferredBlob: BlobEntry | null = null;
      let fallbackBlob:  BlobEntry | null = null;
      for await (const blob of container.listBlobsFlat()) {
        const blobBaseName = blob.name.split('/').pop() ?? blob.name;
        if (!blobBaseName.toLowerCase().endsWith('.xlsx')) continue;
        const modified = blob.properties.lastModified ?? new Date(0);
        const entry: BlobEntry = { name: blob.name, lastModified: modified, etag: blob.properties.etag ?? '' };
        if (blobBaseName.toLowerCase().startsWith('ace-1_combined_radet')) {
          if (!preferredBlob || modified > preferredBlob.lastModified) preferredBlob = entry;
        } else {
          if (!fallbackBlob || modified > fallbackBlob.lastModified) fallbackBlob = entry;
        }
      }
      const latestBlob = preferredBlob ?? fallbackBlob;

      if (!latestBlob) {
        logger.warn(`No .xlsx blobs found in container '${sourceContainer}' â€” falling back to local input/`);
      } else {
        const localName  = latestBlob.name.replace(/[/\\]/g, '_');
        const localPath  = path.join(inputDir, localName);
        const etagCache  = loadEtagCache();
        const cachedEtag = etagCache[latestBlob.name];

        const needsDownload = !fs.existsSync(localPath) || cachedEtag !== latestBlob.etag;

        if (needsDownload) {
          const reason = !fs.existsSync(localPath) ? 'not cached' : 'ETag changed';
          logger.info(`  Downloading: ${latestBlob.name} (${reason}, modified ${latestBlob.lastModified.toISOString().slice(0, 10)})`);
          const blockClient = container.getBlockBlobClient(latestBlob.name);
          const tempPath = `${localPath}.download`;
          if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
          await blockClient.downloadToFile(tempPath);
          if (!hasZipEndMarker(tempPath)) {
            if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
            throw new Error(`Downloaded workbook is incomplete: ${latestBlob.name}`);
          }
          if (fs.existsSync(localPath)) fs.unlinkSync(localPath);
          fs.renameSync(tempPath, localPath);
          etagCache[latestBlob.name] = latestBlob.etag;
          saveEtagCache(etagCache);
          logger.info(`  Saved to: ${localPath}`);
        } else {
          if (!hasZipEndMarker(localPath)) {
            throw new Error(`Cached workbook is incomplete: ${localPath}`);
          }
          logger.info(`  Up to date (ETag unchanged): ${localPath}`);
        }
        return localPath;
      }
    } catch (err) {
      logger.warn(`  Blob source fetch failed: ${String(err)} â€” falling back to local input/`);
    }
  }

  // Local fallback. Downloaded blobs are saved as
  // YYYY-MM-DD_ACE-1_Combined_RADET-YYYY-MM-DD.xlsx, so match by containing
  // the RADET name instead of only startsWith().
  const matches = fs.readdirSync(inputDir)
    .filter(f => f.includes('ACE-1_Combined_RADET') && f.endsWith('.xlsx'))
    .map(f => {
      const fullPath = path.join(inputDir, f);
      const usable = hasZipEndMarker(fullPath);
      if (!usable) logger.warn(`  Skipping incomplete local workbook: ${fullPath}`);
      return { file: f, mtimeMs: fs.statSync(fullPath).mtimeMs, usable };
    })
    .filter(entry => entry.usable)
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  return matches.length > 0 ? path.join(inputDir, matches[0].file) : path.join(inputDir, 'RADET.xlsx');
}

const configDir = path.resolve(process.cwd(), 'config');

async function resolveTargetWorkbookPath(): Promise<string | null> {
  if (process.env['TARGET_FILE']) {
    const targetPath = path.resolve(process.env['TARGET_FILE']);
    return fs.existsSync(targetPath) ? targetPath : null;
  }

  const targetBlobName = process.env['AZURE_TARGET_BLOB'] ?? 'ACE-1_Targets.xlsx';
  const targetContainer =
    process.env['AZURE_TARGET_CONTAINER'] ??
    process.env['AZURE_STORAGE_CONTAINER'] ??
    'powerbi-datasource';
  const localPath = path.join(inputDir, targetBlobName.replace(/[/\\]/g, '_'));
  const connectionString = process.env['AZURE_STORAGE_CONNECTION_STRING'];

  if (connectionString) {
    try {
      logger.info(`Checking target workbook from blob storage (${targetContainer}/${targetBlobName})...`);
      const blobService = BlobServiceClient.fromConnectionString(connectionString);
      const container = blobService.getContainerClient(targetContainer);
      const blobClient = container.getBlockBlobClient(targetBlobName);
      const tempPath = `${localPath}.download`;
      if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
      await blobClient.downloadToFile(tempPath);
      fs.renameSync(tempPath, localPath);
      logger.info(`  Target workbook saved to: ${localPath}`);
      return localPath;
    } catch (err) {
      const tempPath = `${localPath}.download`;
      if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
      logger.warn(`  Target workbook fetch failed: ${String(err)} â€” falling back to local input/`);
    }
  }

  return fs.existsSync(localPath) ? localPath : null;
}

// â”€â”€â”€ Main â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

async function main(): Promise<void> {
  const workbookPath = await resolveWorkbookPath();
  const targetWorkbookPath = await resolveTargetWorkbookPath();

  logger.info('================================================================');
  logger.info('  RADET Dashboard Aggregation Engine  v1.0.0');
  logger.info('================================================================');
  logger.info(`Date mode:      ${dateModeConfig.mode}`);
  logger.info(`Workbook:       ${workbookPath}`);
  logger.info(`Config dir:     ${configDir}`);
  logger.info(`Output dir:     ${outputDir}`);

  if (!fs.existsSync(workbookPath)) {
    logger.error(`Workbook not found: ${workbookPath}`);
    logger.error('Place an ACE-1_Combined_RADET*.xlsx file in the input/ folder, or set RADET_FILE env var.');
    process.exit(1);
  }

  // â”€â”€ Aggregation â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  logger.info('Starting aggregation engine setup...');
  const engine = new AggregationEngine({
    workbookPath,
    configDir,
    dateModeConfig,
    logger,
  });

  await engine.init();
  logger.info('Aggregation engine initialised. Starting workbook processing...');

  // â”€â”€ Output â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  // Stream rows directly to CSV as each indicator finalises â€” never accumulate
  // all rows in memory (avoids OOM with 800K+ rows across 93 indicators).
  const writer = new OutputWriter({ outputDir, logger });
  const targetService = new TargetService(logger);
  if (targetWorkbookPath) {
    try {
      await targetService.load(targetWorkbookPath, configDir);
    } catch (err) {
      logger.warn(`Target workbook could not be loaded (${String(err)}) â€” Target column will remain blank.`);
    }
  } else {
    logger.warn('Target workbook not found â€” DashboardSummary.csv Target column will remain blank.');
  }
  const csvPath = path.join(outputDir, 'DashboardSummary.csv');
  const targetCsvPath = path.join(outputDir, 'DashboardTargets.csv');
  const csvStream = writer.openCsvStream(csvPath);
  const targetCsvStream = writer.openTargetCsvStream(targetCsvPath);
  const emittedTargetKeys = new Set<string>();
  let rowCount = 0;
  let targetRowCount = 0;

  await engine.process((row) => {
    const target = targetService.getTarget(row);
    const targetKey = [row.Period, row.DATIMCode, row.Indicator, row.Sex].join('|');

    if (target && !emittedTargetKeys.has(targetKey)) {
      emittedTargetKeys.add(targetKey);
      writer.writeTargetCsvRow(targetCsvStream, {
        Period: row.Period,
        TargetDate: target.targetDate,
        State: row.State,
        Facility: row.Facility,
        DATIMCode: row.DATIMCode,
        Indicator: row.Indicator,
        Sex: row.Sex,
        Target: target.target,
      });
      targetRowCount++;
    }

    writer.writeCsvRow(csvStream, row);
    rowCount++;
  });

  // ── EDCT data from PostgreSQL (Dashboard sheet only) ─────────────────────
  // Rows are appended to DashboardSummary.csv but NOT passed to the
  // quarterly/monthly workbook sheets (edctRows stays empty below).
  const pgConnString = process.env['POSTGRES_CONNECTION_STRING'];
  if (pgConnString) {
    const edctService = new EdctService(pgConnString, logger);
    const connected   = await edctService.testConnection();
    if (connected) {
      try {
        const edctSqlPath  = path.resolve(process.cwd(), 'scripts', 'edct_query.sql');
        const edctDateRange = new DateHelper(dateModeConfig).getGroupByDateRange();
        const edctRows: DashboardRow[] = await edctService.fetchDashboardRows(edctSqlPath, edctDateRange);
        for (const row of edctRows) {
          writer.writeCsvRow(csvStream, row);
          rowCount++;
        }
        logger.info(`  EDCT rows appended to Dashboard: ${edctRows.length}`);
      } catch (err) {
        logger.error(`  EDCT query failed: ${String(err)}`);
      }
    }
  } else {
    logger.warn('  POSTGRES_CONNECTION_STRING not set — EDCT data skipped.');
  }

  logger.info('Aggregation complete. Closing output CSV streams...');
  await writer.closeCsvStream(csvStream);
  await writer.closeCsvStream(targetCsvStream);

  let dashboardWorkbookPath: string | null = null;
  if (rowCount > 0 && targetWorkbookPath) {
    const workbookWriter = new WorkbookReportWriter(logger);
    dashboardWorkbookPath = await workbookWriter.writeFyDashboardWorkbook({
      outputDir,
      dashboardCsvPath: csvPath,
      targetWorkbookPath,
    });
  }

  if (rowCount > 0) {
    logger.info(`\nOutput files:`);
    logger.info(`  ${csvPath} (${rowCount} rows)`);
    logger.info(`  ${targetCsvPath} (${targetRowCount} rows)`);
    if (dashboardWorkbookPath) logger.info(`  ${dashboardWorkbookPath}`);
  } else {
    logger.warn('No dashboard rows generated — check your workbook and date mode.');
  }

  // â”€â”€ Azure Blob Upload â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const connectionString = process.env['AZURE_STORAGE_CONNECTION_STRING'];
  const containerName    = process.env['AZURE_STORAGE_CONTAINER'] ?? 'powerbi-datasource';
  const blobPrefix       = process.env['AZURE_STORAGE_BLOB_PREFIX'] ?? '';

  async function uploadFileToBlob(filePath: string, contentType: string): Promise<void> {
    const blobServiceClient = BlobServiceClient.fromConnectionString(connectionString!);
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blobName = blobPrefix + path.basename(filePath);
    const blockBlobClient = containerClient.getBlockBlobClient(blobName);
    const fileSize = fs.statSync(filePath).size;
    await blockBlobClient.uploadStream(fs.createReadStream(filePath), undefined, undefined, {
      blobHTTPHeaders: { blobContentType: contentType },
    });
    logger.info(`  Uploaded: ${blobName} (${(fileSize / 1024).toFixed(1)} KB) -> ${containerName}`);
  }

  // ── Validation Report ─────────────────────────────────────────────────────
  let validationReportPath: string | null = null;
  if (engine.validationIssues.length > 0) {
    validationReportPath = await writer.writeValidationReport(engine.validationIssues);
    logger.warn(`\nValidation issues found. Report: ${validationReportPath}`);
  }

  // ── Upload CSV + XLSX together ─────────────────────────────────────────────
  if (rowCount > 0 && connectionString) {
    const filesToUpload: Array<{ filePath: string; contentType: string }> = [];
    if (fs.existsSync(csvPath)) {
      filesToUpload.push({ filePath: csvPath, contentType: 'text/csv' });
    }
    if (dashboardWorkbookPath && fs.existsSync(dashboardWorkbookPath)) {
      filesToUpload.push({ filePath: dashboardWorkbookPath, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }

    logger.info(`\nUploading ${filesToUpload.length} file(s) to Azure Blob Storage (${containerName})...`);
    let allUploaded = true;
    for (const { filePath, contentType } of filesToUpload) {
      try {
        await uploadFileToBlob(filePath, contentType);
      } catch (err) {
        logger.error(`  Upload failed [${path.basename(filePath)}]: ${String(err)}`);
        allUploaded = false;
      }
    }

    // ── Clear outputs after successful upload ──────────────────────────────
    // Only process.log is kept — everything else is deleted to save disk space.
    if (allUploaded) {
      logger.info('\nClearing output files after successful upload...');
      const toDelete = [
        csvPath,
        targetCsvPath,
        dashboardWorkbookPath,
        validationReportPath,
      ].filter((p): p is string => !!p && fs.existsSync(p));

      for (const filePath of toDelete) {
        try {
          fs.unlinkSync(filePath);
          logger.info(`  Deleted: ${path.basename(filePath)}`);
        } catch (err) {
          logger.warn(`  Could not delete ${path.basename(filePath)}: ${String(err)}`);
        }
      }
    } else {
      logger.warn('  Some uploads failed — output files kept for inspection.');
    }
  } else if (rowCount > 0 && !connectionString) {
    logger.warn('  AZURE_STORAGE_CONNECTION_STRING not set — skipping blob upload.');
  }

  // â”€â”€ Final Stats â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  logger.info('\nâ”€â”€â”€ Processing Summary â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€');
  logger.info(`Rows read:             ${engine.stats.totalRows}`);
  logger.info(`Rows after filtering:  ${engine.stats.filteredRows}`);
  logger.info(`Dashboard rows out:    ${engine.stats.aggregatedRows}`);
  logger.info(`Duration:              ${engine.stats.durationMs}ms`);
  logger.info(`Indicators done:       ${engine.stats.indicatorsProcessed.join(', ')}`);
  logger.info('â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€\n');
}

main().catch((err) => {
  logger.error(`Fatal error: ${String(err)}`);
  process.exit(1);
});


