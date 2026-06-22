'use strict';

/**
 * service_runner.js
 * Runs the RADET Dashboard Aggregation Engine on a fixed daily schedule.
 * Invoked by NSSM as a long-running Windows Service process.
 *
 * Default schedule: 09:00, 11:00, 13:00, 15:00, 17:00, 19:00, 21:00
 * Override via .env:  RUN_SCHEDULE=9,11,13,15,17,19,21
 */

const { spawn } = require('child_process');
const path = require('path');
const fs   = require('fs');

// ── Load .env ────────────────────────────────────────────────────────────────
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

// ── Schedule config ──────────────────────────────────────────────────────────
// Parse RUN_SCHEDULE from .env, e.g. "9,11,13,15,17,19,21"
// Falls back to default: every 2 hours from 09:00 to 21:00
const DEFAULT_SCHEDULE = [9, 11, 13, 15, 17, 19, 21];

function parseSchedule(raw) {
  if (!raw) return DEFAULT_SCHEDULE;
  const parsed = raw.split(',')
    .map(s => parseInt(s.trim(), 10))
    .filter(n => !isNaN(n) && n >= 0 && n <= 23);
  return parsed.length > 0 ? [...new Set(parsed)].sort((a, b) => a - b) : DEFAULT_SCHEDULE;
}

const RUN_HOURS = parseSchedule(process.env['RUN_SCHEDULE']);
const RUN_ON_START = String(process.env['RUN_ON_START'] ?? 'true').toLowerCase() !== 'false';

// ── Helpers ──────────────────────────────────────────────────────────────────
function timestamp() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ` +
    `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
}

function log(msg) {
  console.log(`[${timestamp()}] ${msg}`);
}

/**
 * Returns the next Date object matching one of RUN_HOURS.
 * If all today's slots have passed, returns the first slot tomorrow.
 */
function getNextRunTime() {
  const now = new Date();

  for (const hour of RUN_HOURS) {
    const candidate = new Date(now);
    candidate.setHours(hour, 0, 0, 0);
    if (candidate > now) return candidate;
  }

  // All slots for today are past — first slot tomorrow
  const tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(RUN_HOURS[0], 0, 0, 0);
  return tomorrow;
}

function msUntil(target) {
  return Math.max(0, target.getTime() - Date.now());
}

function hhmm(date) {
  return date.toTimeString().slice(0, 5);
}

// ── Engine runner ─────────────────────────────────────────────────────────────
function runEngine() {
  return new Promise((resolve) => {
    log('>>> Engine run starting...');
    const proc = spawn(
      process.execPath,
      ['--max-old-space-size=8192', path.join(__dirname, 'dist', 'index.js')],
      { cwd: __dirname, stdio: 'inherit', env: process.env },
    );
    proc.on('error', (err) => {
      log(`Engine failed to start: ${err.message}`);
      resolve(false);
    });
    proc.on('close', (code) => {
      if (code === 0) {
        log('<<< Engine run completed successfully.');
      } else {
        log(`<<< Engine exited with code ${code}.`);
      }
      resolve(code === 0);
    });
  });
}

// ── Main loop ─────────────────────────────────────────────────────────────────
async function main() {
  log('=======================================================');
  log('  RADET Dashboard Aggregation Engine  -  Service Runner');
  log('=======================================================');
  log(`Daily schedule: ${RUN_HOURS.map(h => String(h).padStart(2,'0') + ':00').join('  ')}`);
  log(`Run on service start/restart: ${RUN_ON_START ? 'yes' : 'no'}`);

  if (RUN_ON_START) {
    log('Service just started/restarted; running engine immediately once.');
    await runEngine();
  }

  while (true) {
    const next = getNextRunTime();
    const wait = msUntil(next);
    const waitMins = Math.round(wait / 60000);

    log(`Next run: ${next.toDateString()} at ${hhmm(next)}  (in ${waitMins} min)`);

    // Sleep until next scheduled slot
    await new Promise(resolve => setTimeout(resolve, wait));

    // Double-check we are in a valid run slot (guard against clock drift / DST)
    const nowHour = new Date().getHours();
    if (!RUN_HOURS.includes(nowHour)) {
      log(`Slot check: current hour ${nowHour} not in schedule — skipping.`);
      continue;
    }

    await runEngine();
  }
}

main().catch(err => {
  console.error(`[${timestamp()}] Fatal: ${err.message}`);
  process.exit(1);
});
