'use strict';

/**
 * service_runner.js
 * Runs the RADET Dashboard Aggregation Engine on a fixed daily schedule.
 * Invoked by NSSM as a long-running Windows Service process.
 *
 * Schedule format: "HH:MM,HH:MM,..."  e.g. "8:40,10:40,12:40,14:40,16:40,18:40,20:40"
 * Plain hours also accepted:           e.g. "9,11,13" → treated as 09:00, 11:00, 13:00
 * Override via .env:  RUN_SCHEDULE=8:40,10:40,12:40,14:40,16:40,18:40,20:40
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
// Each slot is { hour: number, minute: number }
const DEFAULT_SCHEDULE = [
  { hour: 8,  minute: 40 },
  { hour: 10, minute: 40 },
  { hour: 12, minute: 40 },
  { hour: 14, minute: 40 },
  { hour: 16, minute: 40 },
  { hour: 18, minute: 40 },
  { hour: 20, minute: 40 },
];

function parseSchedule(raw) {
  if (!raw) return DEFAULT_SCHEDULE;

  const slots = raw.split(',').map(s => {
    const trimmed = s.trim();
    if (trimmed.includes(':')) {
      const [h, m] = trimmed.split(':').map(n => parseInt(n, 10));
      if (!isNaN(h) && !isNaN(m) && h >= 0 && h <= 23 && m >= 0 && m <= 59) {
        return { hour: h, minute: m };
      }
    } else {
      const h = parseInt(trimmed, 10);
      if (!isNaN(h) && h >= 0 && h <= 23) return { hour: h, minute: 0 };
    }
    return null;
  }).filter(Boolean);

  if (slots.length === 0) return DEFAULT_SCHEDULE;

  // Deduplicate and sort by (hour, minute)
  const seen = new Set();
  return slots
    .filter(s => { const k = `${s.hour}:${s.minute}`; if (seen.has(k)) return false; seen.add(k); return true; })
    .sort((a, b) => a.hour !== b.hour ? a.hour - b.hour : a.minute - b.minute);
}

const RUN_SLOTS  = parseSchedule(process.env['RUN_SCHEDULE']);
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

function slotLabel(slot) {
  return `${String(slot.hour).padStart(2, '0')}:${String(slot.minute).padStart(2, '0')}`;
}

/**
 * Returns the next Date matching one of RUN_SLOTS.
 * If all today's slots have passed, returns the first slot tomorrow.
 */
function getNextRunTime() {
  const now = new Date();

  for (const slot of RUN_SLOTS) {
    const candidate = new Date(now);
    candidate.setHours(slot.hour, slot.minute, 0, 0);
    if (candidate > now) return { date: candidate, slot };
  }

  // All slots for today passed — first slot tomorrow
  const tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const first = RUN_SLOTS[0];
  tomorrow.setHours(first.hour, first.minute, 0, 0);
  return { date: tomorrow, slot: first };
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
  log(`Daily schedule: ${RUN_SLOTS.map(slotLabel).join('  ')}`);
  log(`Run on service start/restart: ${RUN_ON_START ? 'yes' : 'no'}`);

  if (RUN_ON_START) {
    log('Service just started/restarted; running engine immediately once.');
    await runEngine();
  }

  while (true) {
    const { date: next, slot } = getNextRunTime();
    const wait = msUntil(next);
    const waitMins = Math.round(wait / 60000);

    log(`Next run: ${next.toDateString()} at ${hhmm(next)}  (in ${waitMins} min)`);

    await new Promise(resolve => setTimeout(resolve, wait));

    // Guard: confirm we are within 2 minutes of the expected slot
    const now = new Date();
    const diffMs = Math.abs(now - new Date(now.getFullYear(), now.getMonth(), now.getDate(), slot.hour, slot.minute, 0, 0));
    if (diffMs > 2 * 60 * 1000) {
      log(`Slot check: woke at ${hhmm(now)} but expected ${slotLabel(slot)} — skipping.`);
      continue;
    }

    await runEngine();
  }
}

main().catch(err => {
  console.error(`[${timestamp()}] Fatal: ${err.message}`);
  process.exit(1);
});
