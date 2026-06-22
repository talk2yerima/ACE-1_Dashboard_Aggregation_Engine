'use strict';

/**
 * service_runner.js
 * Runs the RADET Dashboard Aggregation Engine on a fixed daily schedule.
 * Invoked by NSSM as a long-running Windows Service process.
 *
 * Default schedule: 09:10, 11:10, 13:10, 15:10, 17:10, 19:10, 21:10
 * Override via .env:  RUN_SCHEDULE=9:10,11:10,13:10,15:10,17:10,19:10,21:10
 * Legacy hour-only format still accepted: RUN_SCHEDULE=9,11,13
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
// Parse RUN_SCHEDULE from .env, e.g. "9:10,11:10,13:10,15:10,17:10,19:10,21:10"
// Legacy hour-only format (e.g. "9,11,13") is still accepted and treated as HH:00.
// Falls back to default: every 2 hours from 09:10 to 21:10
const DEFAULT_SCHEDULE = [
  { h: 9,  m: 10 },
  { h: 11, m: 10 },
  { h: 13, m: 10 },
  { h: 15, m: 10 },
  { h: 17, m: 10 },
  { h: 19, m: 10 },
  { h: 21, m: 10 },
];

function parseSchedule(raw) {
  if (!raw) return DEFAULT_SCHEDULE;
  const parsed = raw.split(',')
    .map(s => {
      const parts = s.trim().split(':');
      const h = parseInt(parts[0], 10);
      const m = parts.length > 1 ? parseInt(parts[1], 10) : 0;
      return { h, m };
    })
    .filter(({ h, m }) => !isNaN(h) && !isNaN(m) && h >= 0 && h <= 23 && m >= 0 && m <= 59);
  if (parsed.length === 0) return DEFAULT_SCHEDULE;
  // Deduplicate and sort by time-of-day
  const seen = new Set();
  return parsed
    .filter(({ h, m }) => { const k = `${h}:${m}`; if (seen.has(k)) return false; seen.add(k); return true; })
    .sort((a, b) => a.h !== b.h ? a.h - b.h : a.m - b.m);
}

const RUN_SLOTS = parseSchedule(process.env['RUN_SCHEDULE']);
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
 * Returns the next Date object matching one of RUN_SLOTS.
 * If all today's slots have passed, returns the first slot tomorrow.
 */
function getNextRunTime() {
  const now = new Date();

  for (const { h, m } of RUN_SLOTS) {
    const candidate = new Date(now);
    candidate.setHours(h, m, 0, 0);
    if (candidate > now) return candidate;
  }

  // All slots for today are past — first slot tomorrow
  const tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(RUN_SLOTS[0].h, RUN_SLOTS[0].m, 0, 0);
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
  log(`Daily schedule: ${RUN_SLOTS.map(({ h, m }) => String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0')).join('  ')}`);
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
    // Allow ±2 minutes around the scheduled minute to handle setTimeout jitter.
    const nowDate = new Date();
    const nowHour = nowDate.getHours();
    const nowMin  = nowDate.getMinutes();
    const inSlot  = RUN_SLOTS.some(({ h, m }) => h === nowHour && Math.abs(m - nowMin) <= 2);
    if (!inSlot) {
      log(`Slot check: ${String(nowHour).padStart(2,'0')}:${String(nowMin).padStart(2,'0')} not in schedule — skipping.`);
      continue;
    }

    await runEngine();
  }
}

main().catch(err => {
  console.error(`[${timestamp()}] Fatal: ${err.message}`);
  process.exit(1);
});
