#!/usr/bin/env node
/*
 * Builds data.json from every statement in /reports (xlsx) and /manual (json).
 * Run by the GitHub Action on every push. Run locally with:  node scripts/build.mjs
 *
 * Exits with an error (and publishes nothing) when a file cannot be read,
 * when columns are missing, or when two files cover the same month.
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const XLSX = require(path.join(ROOT, 'vendor/xlsx.full.min.js'));
const Parser = require(path.join(ROOT, 'js/parser.js'));

const errors = [], warnings = [];
const readJson = (p, fallback) => (fs.existsSync(p) ? JSON.parse(fs.readFileSync(p, 'utf8')) : fallback);
const list = (dir) => (fs.existsSync(dir) ? fs.readdirSync(dir).sort() : []);

// ---- config ---------------------------------------------------------------
const config = readJson(path.join(ROOT, 'config.json'), null);
if (!config) { console.error('config.json is missing.'); process.exit(1); }
const weightSum = config.partners.reduce((s, p) => s + p.weight, 0);
config.partners.forEach((p) => { p.share = p.weight / weightSum; });

// ---- statements from /reports -----------------------------------------------
const byMonth = new Map();
for (const name of list(path.join(ROOT, 'reports'))) {
  if (name.startsWith('.') || name.startsWith('~$') || /^readme/i.test(name)) continue;
  if (!/\.xlsx$/i.test(name)) { warnings.push(`reports/${name}: not an .xlsx file, ignored.`); continue; }
  try {
    const wb = XLSX.read(fs.readFileSync(path.join(ROOT, 'reports', name)), Parser.READ_OPTIONS);
    const m = Parser.parseWorkbook(XLSX, wb, name);
    if (byMonth.has(m.month)) {
      errors.push(`Two files cover ${m.month}: "reports/${byMonth.get(m.month).source}" and "reports/${name}". Delete the one you do not want.`);
      continue;
    }
    m.warnings.forEach((w) => warnings.push(`reports/${name}: ${w}`));
    byMonth.set(m.month, m);
  } catch (e) {
    errors.push(`reports/${name}: ${e.message}`);
  }
}

// ---- hand-entered months from /manual (these win over /reports) --------------
for (const name of list(path.join(ROOT, 'manual'))) {
  if (!/\.json$/i.test(name)) continue;
  try {
    const raw = JSON.parse(fs.readFileSync(path.join(ROOT, 'manual', name), 'utf8'));
    if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(raw.month || '')) throw new Error('"month" must look like "2026-04".');
    if (!raw.platforms || !Object.keys(raw.platforms).length) throw new Error('"platforms" is missing.');
    if (byMonth.has(raw.month)) warnings.push(`manual/${name} replaces reports/${byMonth.get(raw.month).source} for ${raw.month}.`);
    const m = Parser.finalize({
      month: raw.month, source: `manual/${name}`, manual: true, rates: raw.rates || {},
      platforms: raw.platforms, statedTotal: raw.statedTotal ?? null, warnings: [],
    });
    byMonth.set(m.month, m);
  } catch (e) {
    errors.push(`manual/${name}: ${e.message}`);
  }
}

if (errors.length) fail();

// ---- month-end EUR to ILS rates ---------------------------------------------
// ECB reference rate for the last business day of each month, fetched once and
// then kept in fx.json so historical shekel figures never change.
const fxPath = path.join(ROOT, 'fx.json');
const fx = readJson(fxPath, {});
const today = new Date().toISOString().slice(0, 10);
const months = [...byMonth.values()].sort((a, b) => a.month.localeCompare(b.month));

for (const m of months) {
  if (fx[m.month]) { m.fx = fx[m.month]; continue; }
  const [y, mo] = m.month.split('-').map(Number);
  const monthEnd = new Date(Date.UTC(y, mo, 0)).toISOString().slice(0, 10);
  try {
    const res = await fetch(`https://api.frankfurter.dev/v1/${monthEnd}?base=EUR&symbols=ILS`, { signal: AbortSignal.timeout(15000) });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const body = await res.json();
    const rate = body?.rates?.ILS;
    if (typeof rate !== 'number') throw new Error('no ILS rate in the response');
    const entry = { rate, date: body.date, source: 'ECB' };
    if (today > monthEnd) fx[m.month] = entry; else entry.provisional = true; // month not over yet: do not freeze
    m.fx = entry;
  } catch (e) {
    m.fx = null;
    warnings.push(`Could not fetch the EUR/ILS rate for ${m.month} (${e.message}). Shekel figures for that month use the nearest known rate until the next run.`);
  }
}
fs.writeFileSync(fxPath, JSON.stringify(Object.fromEntries(Object.entries(fx).sort()), null, 2) + '\n');

// ---- write data.json (keep the timestamp when nothing changed) ---------------
const dataPath = path.join(ROOT, 'data.json');
const previous = readJson(dataPath, null);
const payload = { config, months };
const unchanged = previous && JSON.stringify({ config: previous.config, months: previous.months }) === JSON.stringify(payload);
const out = { generatedAt: unchanged ? previous.generatedAt : new Date().toISOString(), ...payload };
fs.writeFileSync(dataPath, JSON.stringify(out, null, 2) + '\n');

// ---- report -------------------------------------------------------------------
const lines = months.map((m) => {
  const flags = m.checks.length ? `  (${m.checks.length} line(s) differ from the rates)` : '';
  return `  ${m.month}  owner payment €${m.ownerPayment.toFixed(2)}  ${m.fx ? '1 EUR = ' + m.fx.rate + ' ILS' : 'no rate'}  ${m.source}${flags}`;
});
console.log(`Built data.json with ${months.length} month(s):\n${lines.join('\n')}`);
warnings.forEach((w) => console.log(`${process.env.GITHUB_ACTIONS ? '::warning::' : 'Warning: '}${w}`));
summary(`### Dashboard data built\n\n${months.length} month(s) published.\n\n` + (warnings.length ? '**Warnings**\n\n' + warnings.map((w) => `- ${w}`).join('\n') : ''));

function fail() {
  console.error('\nThe dashboard was NOT updated. Fix the following and upload again:\n');
  errors.forEach((e) => console.error(`${process.env.GITHUB_ACTIONS ? '::error::' : '  - '}${e}`));
  summary('### The dashboard was not updated\n\n' + errors.map((e) => `- ${e}`).join('\n'));
  process.exit(1);
}
function summary(md) {
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, md + '\n');
}
