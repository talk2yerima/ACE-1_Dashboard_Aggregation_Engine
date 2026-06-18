/**
 * service_runner.js
 * Runs the RADET Dashboard Aggregation Engine on a configurable schedule.
 * Invoked by NSSM as a long-running Windows Service process.
 *
 * Config (via .env or environment):
 *   RUN_INTERVAL_HOURS  - Hours between engine runs (default: 6)
 *   DATE_MODE           - Passed through to the engine
 */

'use strict';

const { spawn } = require('child_process');
const path = require('path');
const fs   = require('fs');

// Load .env if present
const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq < 0) continue;
    const key = trimmed.slice(0, eq).trim();
    const val = trimmed.slice(eq + 1).trim();
    if (key && !(key in process.env)) process.env[key] = val;
  }
}

const INTERVAL_HOURS = parseFloat(process.env['RUN_INTERVAL_HOURS'] ?? '6');
const INTERVAL_MS    = Math.max(INTERVAL_HOURS * 3600 * 1000, 60_000); // minimum 1 min

function timestamp() {
  return new Date().toISOString().replace('T', ' ').slice(0, 19);
}

function log(msg) {
  console.log(`[${timestamp()}] ${msg}`);
}

function runEngine() {
  return new Promise((resolve) => {
    log('Starting engine run...');
    const proc = spawn(
      process.execPath,           // node.exe
      ['--max-old-space-size=8192', path.join(__dirname, 'dist', 'index.js')],
      { cwd: __dirname, stdio: 'inherit', env: process.env },
    );
    proc.on('error', (err) => {
      log(`Failed to start engine: ${err.message}`);
      resolve(false);
    });
    proc.on('close', (code) => {
      if (code === 0) {
        log('Engine run completed successfully.');
        resolve(true);
      } else {
        log(`Engine exited with code ${code}.`);
        resolve(false);
      }
    });
  });
}

async function main() {
  log('=======================================================');
  log('  RADET Dashboard Aggregation Engine  -  Service Runner');
  log('=======================================================');
  log(`Interval: every ${INTERVAL_HOURS} hour(s)`);

  // Run immediately on startup, then repeat on interval
  while (true) {
    await runEngine();
    const nextRun = new Date(Date.now() + INTERVAL_MS);
    log(`Next run scheduled at ${nextRun.toISOString().replace('T', ' ').slice(0, 19)}`);
    await new Promise(resolve => setTimeout(resolve, INTERVAL_MS));
  }
}

main().catch(err => {
  console.error(`[${timestamp()}] Fatal: ${err.message}`);
  process.exit(1);
});
