import 'dotenv/config';
import path from 'path';
import fs from 'fs';
import winston from 'winston';
import { BlobServiceClient } from '@azure/storage-blob';

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
fs.mkdirSync(inputDir, { recursive: true });

// ── Resolve workbook path ──────────────────────────────────────────────────
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
        logger.warn(`No .xlsx blobs found in container '${sourceContainer}' — falling back to local input/`);
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
          await blockClient.downloadToFile(localPath);
          etagCache[latestBlob.name] = latestBlob.etag;
          saveEtagCache(etagCache);
          logger.info(`  Saved to: ${localPath}`);
        } else {
          logger.info(`  Up to date (ETag unchanged): ${localPath}`);
        }
        return localPath;
      }
    } catch (err) {
      logger.warn(`  Blob source fetch failed: ${String(err)} — falling back to local input/`);
    }
  }

  // Local fallback
  const match = fs.readdirSync(inputDir).find(f => f.startsWith('ACE-1_Combined_RADET') && f.endsWith('.xlsx'));
  return match ? path.join(inputDir, match) : path.join(inputDir, 'RADET.xlsx');
}

const configDir = path.resolve(process.cwd(), 'config');

// ─── Main ────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  const workbookPath = await resolveWorkbookPath();

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
  let csvPath: string | null = null;

  if (dashboardRows.length > 0) {
    const { xlsx, csv } = await writer.writeDashboard(dashboardRows);
    csvPath = csv;
    logger.info(`\nOutput files:`);
    logger.info(`  ${xlsx}`);
    logger.info(`  ${csv}`);
  } else {
    logger.warn('No dashboard rows generated — check your workbook and date mode.');
  }

  // ── Azure Blob Upload ──────────────────────────────────────────────────
  const connectionString = process.env['AZURE_STORAGE_CONNECTION_STRING'];
  const containerName    = process.env['AZURE_STORAGE_CONTAINER'] ?? 'powerbi-datasource';
  const blobPrefix       = process.env['AZURE_STORAGE_BLOB_PREFIX'] ?? '';

  if (csvPath && connectionString) {
    try {
      logger.info(`\nUploading CSV to Azure Blob Storage (${containerName})...`);
      const blobServiceClient = BlobServiceClient.fromConnectionString(connectionString);
      const containerClient   = blobServiceClient.getContainerClient(containerName);
      const blobName          = blobPrefix + path.basename(csvPath);
      const blockBlobClient   = containerClient.getBlockBlobClient(blobName);
      const fileSize          = fs.statSync(csvPath).size;
      await blockBlobClient.uploadStream(fs.createReadStream(csvPath), undefined, undefined, {
        blobHTTPHeaders: { blobContentType: 'text/csv' },
      });
      logger.info(`  Uploaded: ${blobName} (${(fileSize / 1024).toFixed(1)} KB) → ${containerName}`);
    } catch (err) {
      logger.error(`  Blob upload failed: ${String(err)}`);
    }
  } else if (csvPath && !connectionString) {
    logger.warn('  AZURE_STORAGE_CONNECTION_STRING not set — skipping blob upload.');
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
