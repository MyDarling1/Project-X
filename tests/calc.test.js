#!/usr/bin/env node
/**
 * Unit tests for the pure calculation core of Uzum PVZ dashboard.
 * These functions are extracted from index.html <script> for testability.
 * Run: node tests/calc.test.js
 * 
 * To keep in sync: the PURE FUNCTIONS block below should match the corresponding
 * logic in index.html (after the data guard and DEFAULTS).
 */

const fs = require('fs');
const path = require('path');

// Load real data for realistic tests
const dataPath = path.join(__dirname, '..', 'data.json');
const raw = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const DATA = raw.DATA;
const WK = raw.WK || {};
const FACT = raw.FACT || { months: {} };
const SAL = raw.SAL || {};

// Replicate minimal state and pure functions from dashboard
const DEFAULTS = { fx: 12000, chk: 14, uk: 320 };
let S = { fx: DEFAULTS.fx, chk: DEFAULTS.chk, uk: DEFAULTS.uk, boost: 1, flat: null, scen: 'guar', seg: 'all' };

const segColor = {1:'var(--s1)',2:'var(--s2)',3:'var(--s3)'}; // not used in pure calc

function compute(d, ordersOverride) {
  const orders = (ordersOverride != null) ? ordersOverride : d.runRate;
  const margin = S.chk * d.comm;
  const cost = d.rentSum + d.salarySum + (d.miscSum || 0) + S.uk * S.fx;
  const be = (cost / S.fx) / margin / 30;
  const rev = orders * S.chk * S.fx;
  const commRev = orders * margin * 30 * S.fx;
  const reward = Math.max(d.guarFull, commRev);
  const pnl = reward - cost;
  return { orders, margin, be, rev, commRev, cost, reward, pnl, health: orders / be, gap: orders - be };
}

const hasFact = (typeof FACT !== 'undefined') && FACT && FACT.months && Object.keys(FACT.months).length > 0;
const factMonths = hasFact ? FACT.months : {};
const salFact = (code, key) => { const m = (typeof SAL !== 'undefined' && SAL) ? SAL[key] : null; return (m && m[code] != null) ? m[code] : null; };

function daysInMonthKey(key) {
  const p = String(key).split('-'); return new Date(+p[0], +p[1], 0).getDate();
}

function activeDaysOf(code, key) {
  const d = DATA.find(x => x.code === code), dim = daysInMonthKey(key);
  if (!d || !d.start) return dim;
  const p = String(key).split('-'), st = new Date(d.start + 'T00:00:00');
  const mStart = new Date(+p[0], +p[1] - 1, 1), mEnd = new Date(+p[0], +p[1] - 1, dim);
  if (isNaN(st) || st <= mStart) return dim; if (st > mEnd) return 0;
  return dim - st.getDate() + 1;
}

function factOf(d) {
  if (!hasFact) return null;
  const m = factMonths[FACT.latest]; if (!m || m.income[d.code] == null) return null;
  const inc = m.income[d.code], dim = daysInMonthKey(FACT.latest);
  const cost = (d.rentSum + d.salarySum + (d.miscSum || 0) + S.uk * S.fx) * (activeDaysOf(d.code, FACT.latest) / dim);
  return { income: inc, cost, pnl: inc - cost };
}

const beOf = d => (d.rentSum + d.salarySum + (d.miscSum || 0) + S.uk * S.fx) / S.fx / (S.chk * d.comm) / 30;

function baseOrders(d) {
  const m = (typeof monthList !== 'undefined' && monthList.length) ? monthList[monthList.length - 1] : null;
  if (!m) return d.runRate;
  let o = 0, dd = 0;
  // Simplified: use WK if available
  if (WK && WK.pvz && WK.pvz[d.code]) {
    // For test, fall back to runRate approximation
  }
  return d.runRate;
}

function simOf(d) {
  const m = (monthList && monthList.length) ? monthList[monthList.length-1] : null, key = m ? m.key : null, fm = key && factMonths[key];
  const base = d.runRate, orders = (S.flat != null) ? S.flat : base * S.boost;
  const cm = d.rentSum + d.salarySum + (d.miscSum || 0) + S.uk * S.fx;
  const be = beOf(d), rev = orders * S.chk * S.fx;
  if (fm) {
    const DIM = daysInMonthKey(key), sal = salFact(d.code, key);
    const cost = d.rentSum + (d.miscSum || 0) + S.uk * S.fx + (sal != null ? sal : d.salarySum);
    const factInc = fm.income[d.code] || 0, commRev = d.comm * orders * S.chk * S.fx * DIM;
    const reward = S.scen === 'guar' ? Math.max(factInc, commRev) : commRev;
    return { orders, margin: S.chk * d.comm, be, rev, commRev, cost, reward, pnl: reward - cost, health: be > 0 ? orders / be : 0, gap: orders - be };
  }
  const cost = cm, commRev = orders * S.chk * d.comm * 30 * S.fx;
  const reward = S.scen === 'guar' ? Math.max(d.guarFull, commRev) : commRev;
  return { orders, margin: S.chk * d.comm, be, rev, commRev, cost, reward, pnl: reward - cost, health: be > 0 ? orders / be : 0, gap: orders - be };
}

function netGrowth() {
  const f = [];
  (WK.weeks || []).forEach((w, i) => {
    if ((WK.dayCounts && WK.dayCounts[i]) >= 7 && WK.net && WK.net.ord[i] > 0) f.push(WK.net.ord[i]);
  });
  if (f.length < 2) return null;
  let p = 1, n = 0;
  for (let i = 1; i < f.length; i++) { if (f[i-1] > 0) { p *= f[i] / f[i-1]; n++; } }
  return n ? Math.pow(p, 1 / n) : null;
}

// Simple test runner
let passed = 0, failed = 0;

function test(name, fn) {
  try {
    fn();
    console.log(`✓ ${name}`);
    passed++;
  } catch (e) {
    console.error(`✗ ${name}: ${e.message}`);
    failed++;
  }
}

function assertEqual(actual, expected, msg = '') {
  if (actual !== expected && JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${msg} Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}

console.log('Running calc core unit tests...');
console.log(`Loaded ${DATA.length} PVZs, ${WK.weeks ? WK.weeks.length : 0} weeks, FACT months: ${Object.keys(FACT.months || {}).length}`);

// Basic smoke on real data
test('DATA has 18 PVZs', () => {
  assertEqual(DATA.length, 18);
});

test('compute returns sensible BE and PNL for first PVZ', () => {
  const d = DATA[0];
  const res = compute(d);
  if (res.be <= 0) throw new Error('BE should be positive');
  if (typeof res.pnl !== 'number') throw new Error('PNL should be number');
  console.log(`  Sample BE: ${res.be.toFixed(1)}, PNL: ${res.pnl}`);
});

test('beOf is consistent with compute', () => {
  const d = DATA[0];
  const b1 = beOf(d);
  const b2 = compute(d).be;
  assertEqual(Math.round(b1 * 100), Math.round(b2 * 100)); // tolerance for floating
});

test('netGrowth returns number or null', () => {
  const g = netGrowth();
  if (g !== null && typeof g !== 'number') throw new Error('netGrowth should be number or null');
  console.log(`  netGrowth: ${g}`);
});

test('factOf returns null or object with numbers when FACT present', () => {
  const d = DATA[0];
  const f = factOf(d);
  if (f !== null) {
    if (typeof f.income !== 'number' || typeof f.pnl !== 'number') throw new Error('factOf fields wrong');
  }
});

// Test with forced S values
test('compute respects S.chk and S.fx', () => {
  const d = DATA[0];
  S.chk = 14; S.fx = 12000;
  const r1 = compute(d);
  S.chk = 15;
  const r2 = compute(d);
  if (r2.margin <= r1.margin) throw new Error('Higher chk should give higher margin');
  S.chk = DEFAULTS.chk; // reset
});

console.log(`\n${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
console.log('All tests passed!');

// Note: In real usage, sync the PURE FUNCTIONS block with index.html after changes.