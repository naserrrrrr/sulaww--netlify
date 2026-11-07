#!/usr/bin/env node
/**
 * Unit tests for assets/metric_engine.js (data-contract.md §5 semantics).
 *
 * Synthetic in-memory records only (no sample data). The fixture mirrors
 * tests/test_metric_engine.py so both engines are checked against the same hand-computed
 * numbers. A final test runs scripts/metric_engine.py on the same fixture (when python3 is
 * available) and deep-compares every result for many filter combinations.
 *
 * Run:  node tests/test_metric_engine_js.mjs        (exit 0 = all passed)
 */
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SKILL = path.dirname(HERE);
const ENGINE_PATH = path.join(SKILL, 'assets', 'metric_engine.js');
const require = createRequire(import.meta.url);
const lib = require(ENGINE_PATH);
const { createEngine } = lib;

// ---------------------------------------------------------------------------
// Fixture helpers (same records as tests/test_metric_engine.py)
// ---------------------------------------------------------------------------
const KNOWN = ['value', 'derived', 'zero_recorded', 'zero_confirmed'];

function tv(v = null, s = null) {
  if (s === null) s = v === null ? 'blank' : (v === 0 ? 'zero_recorded' : 'value');
  return { v: KNOWN.includes(s) ? v : null, s, raw: v === null ? null : String(v), ref: 'T!A1', note: null };
}

function rec(id, o = {}, fields = {}) {
  const d = Object.assign({ rtype: 'asset', scope: 'in_scope', use: 'investment', inc: 'rent', own: ['owned'],
    atype: 'عمارة', loc: 'جنوب', mfo: false }, o);
  return {
    id, name: id, record_type: d.rtype, scope: d.scope,
    dims: { use: d.use, income_nature: d.inc, ownership: d.own.slice(), waqf: 'unknown', asset_type: d.atype,
      location: d.loc, managed_for_others: d.mfo },
    fields
  };
}

function records() {
  return [
    // A1: investment rent, rev 100 / val 1000 in 2025; rev 80 in 2024; exp 10
    rec('A1', {}, { 'rev.2025': tv(100), 'rev.2024': tv(80), 'val.2024': tv(1000), 'exp.2025': tv(10) }),
    // A2: investment rent, recorded zero revenue (included in ratio), val 500
    rec('A2', {}, { 'rev.2025': tv(0), 'rev.2024': tv(0), 'val.2024': tv(500), 'exp.2025': tv(5) }),
    // A3: revenue blank (unknown) -> num_unknown; val 2000
    rec('A3', { loc: 'شمال' }, { 'rev.2025': tv(null), 'rev.2024': tv(50), 'val.2024': tv(2000) }),
    // A4: revenue 300, valuation zero_as_missing -> den_unknown
    rec('A4', { atype: 'أرض' }, { 'rev.2025': tv(300), 'rev.2024': tv(null), 'val.2024': tv(null, 'zero_as_missing') }),
    // A5: service asset, no revenue fields at all, val 700
    rec('A5', { use: 'service', inc: 'none', loc: 'شمال' }, { 'val.2024': tv(700) }),
    // A6: managed for others
    rec('A6', { own: ['managed_for_others'], mfo: true }, { 'rev.2025': tv(999), 'val.2024': tv(999) }),
    // A7: operating activity (school)
    rec('A7', { rtype: 'operating_activity', use: 'service', inc: 'operating', own: ['usufruct'] }, { 'rev.2025': tv(400) }),
    // A8: isolated (outside base_where)
    rec('A8', { scope: 'isolated' }, { 'rev.2025': tv(5000), 'val.2024': tv(5000) }),
    // A9: revenue confirmed zero, known negative valuation -> den_nonpositive
    rec('A9', {}, { 'rev.2025': tv(0, 'zero_confirmed'), 'val.2024': { v: -1, s: 'value', raw: '-1', ref: 'T!A9' } })
  ];
}

const ASSET = { in: ['record_type', ['asset', 'land_parcel']] };
const NOT_MFO = { not: { eq: ['managed_for_others', true] } };

function config(over = {}) {
  const cfg = {
    contract_version: '1.0',
    periods: [
      { id: '2025', label: '2025', rev_field: 'rev.2025', exp_field: 'exp.2025', prev_id: '2024',
        prev_rev_field: 'rev.2024', comparable_with_prev: { status: 'unverified', reason: 'x' } },
      { id: '2024', label: '2024', rev_field: 'rev.2024', exp_field: null, prev_id: null,
        prev_rev_field: null, comparable_with_prev: { status: 'unverified', reason: 'none' } }
    ],
    valuation: { field: 'val.2024', year: '2024' },
    base_where: { all: [{ eq: ['scope', 'in_scope'] }] },
    metrics: [],
    breakdowns: [],
    trend: { enabled: true, label: 't', fields: [{ period: '2024', field: 'rev.2024' }, { period: '2025', field: 'rev.2025' }],
      where: { all: [{ eq: ['income_nature', 'rent'] }, NOT_MFO] }, comparable: { status: 'unverified', reason: 'u' } },
    filters: ['asset_type', 'location', 'use'],
    parity_max_values_per_dim: 3
  };
  return Object.assign(cfg, over);
}

const engine = (recs = null, over = {}) => createEngine({ records: recs === null ? records() : recs, config: config(over) });
const close = (a, b, tol = 1e-12) => a !== null && b !== null && Math.abs(a - b) <= tol * Math.max(1, Math.abs(a), Math.abs(b));
const ev = lib.evalWhere;

// ---------------------------------------------------------------------------
// Where
// ---------------------------------------------------------------------------
function test_where_ops() {
  const r = records()[0];
  assert.ok(ev({ eq: ['use', 'investment'] }, r));
  assert.ok(ev({ eq: ['id', 'A1'] }, r));
  assert.ok(!ev({ eq: ['missing_dim', 'x'] }, r));
  assert.ok(ev({ eq: ['missing_dim', null] }, r));           // missing attr => null
  assert.ok(ev({ in: ['record_type', ['asset', 'land_parcel']] }, r));
  assert.ok(ev({ has: ['ownership', 'owned'] }, r));
  assert.ok(!ev({ has: ['ownership', 'usufruct'] }, r));
  assert.ok(!ev({ eq: ['ownership', 'owned'] }, r));          // eq never matches a list
  assert.ok(ev({ has: ['use', 'investment'] }, r));           // has on a scalar = strict eq
  assert.ok(ev({ not: { eq: ['managed_for_others', true] } }, r));
  assert.ok(!ev({ eq: ['managed_for_others', 0] }, r));       // strict: false !== 0
  assert.ok(!ev({ eq: ['managed_for_others', null] }, r));    // false !== null
  assert.ok(ev({ known: 'rev.2025' }, r));
  assert.ok(!ev({ known: 'rev.missing' }, r));
  assert.ok(ev({ status: ['rev.2025', ['value']] }, r));
  assert.ok(!ev({ status: ['rev.missing', ['blank']] }, r));  // absent field has status null
  assert.ok(ev({ gt: ['rev.2025', 0] }, r));
  assert.ok(!ev({ gt: ['rev.2025', 100] }, r));
  assert.ok(ev({ all: [] }, r) && !ev({ any: [] }, r));
  assert.ok(ev({ any: [{ eq: ['use', 'x'] }, { eq: ['use', 'investment'] }] }, r));
  assert.ok(ev(null, r));                                     // no where = true
  const r2 = records()[1];
  assert.ok(!ev({ gt: ['rev.2025', 0] }, r2));
  assert.ok(ev({ status: ['rev.2025', ['zero_recorded', 'zero_confirmed']] }, r2));
  // a dim literally named like an Object.prototype member is just missing
  assert.ok(ev({ eq: ['constructor', null] }, r) && ev({ eq: ['toString', null] }, r));
  assert.throws(() => ev({ eq: ['use', 'x'], in: ['use', []] }, r));   // more than one operator
  assert.throws(() => ev({ like: ['use', 'x'] }, r));                  // unknown operator
}

function test_placeholders_in_where() {
  const eng = engine();
  const spec = { id: 'zr', kind: 'sum', field: '{VAL}', unit: 'SAR',
    where: { all: [{ eq: ['use', 'investment'] }, ASSET, NOT_MFO, { status: ['{REV}', ['zero_recorded', 'zero_confirmed']] }] } };
  const res = eng.compute(spec, { period: '2025' });
  // A2 (zero_recorded, val 500) and A9 (zero_confirmed, val -1) -> 499
  assert.deepEqual(res.included_ids, ['A2', 'A9']);
  assert.equal(res.value, 499);
  assert.equal(res.status, 'complete');
  const res24 = eng.compute(spec, { period: '2024' });
  // 2024: only A2 has rev.2024 zero_recorded
  assert.deepEqual(res24.included_ids, ['A2']);
  assert.equal(res24.value, 500);
  assert.equal(eng.resolveField('{REV}', { period: '2024' }), 'rev.2024');
  assert.equal(eng.resolveField('{EXP}', { period: '2024' }), null);
  assert.equal(eng.resolveField('{VAL}'), 'val.2024');
  assert.equal(eng.resolveField('plain.field'), 'plain.field');
}

// ---------------------------------------------------------------------------
// Kinds
// ---------------------------------------------------------------------------
function test_count() {
  const eng = engine();
  let res = eng.compute({ id: 'c', kind: 'count', where: ASSET, unit: 'asset' });
  // in scope assets: A1-A6, A9 = 7 (A7 operating, A8 isolated)
  assert.equal(res.value, 7);
  assert.equal(res.status, 'complete');
  assert.equal(res.n_included, 7);
  assert.equal(res.empty_cause, null);
  assert.equal(res.note, null);
  // nothing qualifies in the whole scope: a count is a real zero (empty_as_zero defaults to true for count)
  res = eng.compute({ id: 'c', kind: 'count', where: { eq: ['record_type', 'unit'] }, unit: 'count' });
  assert.equal(res.value, 0);
  assert.equal(res.status, 'complete');
  assert.equal(res.n_included, 0);
  assert.equal(res.empty_cause, 'scope');
  assert.equal(res.note, lib.EMPTY_SCOPE_REASON);
  assert.equal(res.reason, null);
  res = eng.compute({ id: 'c', kind: 'count', where: { eq: ['record_type', 'unit'] }, unit: 'count', empty_as_zero: false, empty_text: 'لا وحدات' });
  assert.equal(res.status, 'empty');
  assert.equal(res.reason, 'لا وحدات');
  // a dim filter that empties a non-empty scope stays "empty" with cause filters
  res = eng.compute({ id: 'c', kind: 'count', where: ASSET, unit: 'asset' }, { location: 'nowhere' });
  assert.equal(res.status, 'empty');
  assert.equal(res.value, 0);
  assert.equal(res.empty_cause, 'filters');
  assert.equal(res.reason, lib.EMPTY_FILTERS_REASON);
}

function test_sum_statuses() {
  const eng = engine();
  let res = eng.compute({ id: 'pr', kind: 'sum', field: '{REV}', unit: 'SAR',
    where: { all: [{ eq: ['income_nature', 'rent'] }, NOT_MFO] } });
  // R = A1,A2,A3,A4,A9 ; K = A1(100), A2(0), A4(300), A9(0) ; U = A3(blank) -> 400 partial
  assert.equal(res.value, 400);
  assert.equal(res.status, 'partial');
  assert.equal(res.n_scope, 5);
  assert.equal(res.n_included, 4);
  assert.equal(res.n_unknown, 1);
  assert.deepEqual(res.excluded, [{ id: 'A3', reason: 'blank' }]);
  // absent field -> reason "absent", not_computable
  res = eng.compute({ id: 'x', kind: 'sum', field: 'rev.2025', where: { eq: ['id', 'A5'] }, unit: 'SAR' });
  assert.equal(res.status, 'not_computable');
  assert.equal(res.value, null);
  assert.deepEqual(res.excluded, [{ id: 'A5', reason: 'absent' }]);
  assert.equal(res.reason, 'لا توجد قيم معروفة للحقل ضمن النطاق');
  // empty
  res = eng.compute({ id: 'x', kind: 'sum', field: 'rev.2025', where: { eq: ['id', 'ZZ'] }, unit: 'SAR' });
  assert.equal(res.status, 'empty');
  assert.equal(res.value, null);
  assert.equal(res.n_scope, 0);
  // zero_as_missing never summed: 1000 + 500 + 2000 + 700 - 1 = 4199 (A4 excluded, A6 managed)
  res = eng.compute({ id: 'v', kind: 'sum', field: '{VAL}', where: { all: [ASSET, NOT_MFO] }, unit: 'SAR' });
  assert.equal(res.value, 4199);
  assert.equal(res.status, 'partial');
  assert.ok(res.excluded.some((e) => e.id === 'A4' && e.reason === 'zero_as_missing'));
  // complete: derived status counts as known; A1 100 + A2 0 (derived 0 is still a number)
  const recs = [rec('D1', {}, { 'rev.2025': { v: 100, s: 'derived' } }), rec('D2', {}, { 'rev.2025': tv(0, 'zero_confirmed') })];
  res = createEngine({ records: recs, config: config() }).compute({ id: 's', kind: 'sum', field: '{REV}', unit: 'SAR' });
  assert.equal(res.value, 100);
  assert.equal(res.status, 'complete');
  assert.deepEqual(res.excluded, []);
}

function test_unknown_statuses_never_summed() {
  const statuses = ['zero_as_missing', 'blank', 'filler', 'text', 'error', 'not_applicable', 'not_available'];
  const recs = statuses.map((s, i) => rec('U' + i, {}, { 'rev.2025': { v: null, s, raw: 'x' } }));
  recs.push(rec('U9', {}, { 'rev.2025': tv(7) }));
  const res = createEngine({ records: recs, config: config() }).compute({ id: 's', kind: 'sum', field: '{REV}', unit: 'SAR' });
  assert.equal(res.value, 7);
  assert.equal(res.status, 'partial');
  assert.deepEqual(res.excluded.map((e) => e.reason), statuses);
}

function test_ratio_zero_included_and_exclusion_reasons() {
  const eng = engine();
  const spec = { id: 'gy', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, unit: '%',
    where: { all: [{ eq: ['use', 'investment'] }, ASSET, { not: { eq: ['income_nature', 'operating'] } }, NOT_MFO] } };
  let res = eng.compute(spec);
  // R = A1,A2,A3,A4,A9 ; I = A1 (100/1000), A2 (0/500) -> 100/1500*100 = 6.666...
  assert.deepEqual(res.included_ids, ['A1', 'A2']);
  assert.equal(res.num, 100);
  assert.equal(res.den, 1500);
  assert.ok(close(res.value, 100 / 1500 * 100));
  assert.equal(res.status, 'partial');
  assert.deepEqual(res.excluded, [{ id: 'A3', reason: 'num_unknown' }, { id: 'A4', reason: 'den_unknown' },
    { id: 'A9', reason: 'den_nonpositive' }]);
  // num_known_total = A1 100 + A2 0 + A4 300 + A9 0 = 400 ; coverage_num_share = 100/400
  assert.equal(res.num_known_total, 400);
  assert.ok(close(res.coverage_num_share, 0.25));
  res = eng.compute(Object.assign({}, spec, { where: { eq: ['id', 'A1'] } }));
  assert.equal(res.status, 'complete');
  assert.ok(close(res.value, 10));
  res = eng.compute(Object.assign({}, spec, { where: { eq: ['id', 'A2'] } }));
  assert.equal(res.value, 0);                     // zero revenue included, value 0
  assert.equal(res.status, 'complete');
  assert.equal(res.coverage_num_share, null);     // num_known_total 0
  res = eng.compute(Object.assign({}, spec, { where: { eq: ['id', 'A3'] } }));
  assert.equal(res.status, 'not_computable');
  assert.equal(res.value, null);
  assert.equal(res.n_included, 0);
  // scale omitted -> 1 ; scale null -> 1
  res = eng.compute({ id: 'r1', kind: 'ratio', num: '{REV}', den: '{VAL}', where: { eq: ['id', 'A1'] } });
  assert.ok(close(res.value, 0.1));
  res = eng.compute({ id: 'r1', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: null, where: { eq: ['id', 'A1'] } });
  assert.ok(close(res.value, 0.1));
}

function test_diff_and_missing_expense() {
  const eng = engine();
  const spec = { id: 'noi', kind: 'diff', a: '{REV}', b: '{EXP}', unit: 'SAR',
    where: { all: [{ eq: ['income_nature', 'rent'] }, NOT_MFO] }, unavailable_reason: 'لا مصروفات لهذه الفترة' };
  let res = eng.compute(spec, { period: '2025' });
  // I = A1 (100-10), A2 (0-5) -> 85 ; A3 a_unknown ; A4, A9 b_unknown
  assert.equal(res.value, 85);
  assert.equal(res.status, 'partial');
  assert.equal(res.n_included, 2);
  assert.equal(res.a_sum, 100);
  assert.equal(res.b_sum, 15);
  assert.deepEqual(res.excluded[0], { id: 'A3', reason: 'a_unknown' });
  assert.deepEqual(res.excluded.slice(1).map((e) => e.reason), ['b_unknown', 'b_unknown']);
  // exp placeholder null in 2024 -> not_computable with unavailable_reason
  res = eng.compute(spec, { period: '2024' });
  assert.equal(res.status, 'not_computable');
  assert.equal(res.reason, 'لا مصروفات لهذه الفترة');
  assert.equal(res.n_included, 0);
  assert.equal(res.value, null);
  assert.equal(res.unresolved, '{EXP}');
  // field defined but no record of R has the key -> «لا توجد بيانات مصروفات»
  res = eng.compute(Object.assign({}, spec, { b: 'exp.none' }), { period: '2025' });
  assert.equal(res.status, 'not_computable');
  assert.equal(res.reason, lib.NO_EXPENSE_REASON);
  assert.equal(res.value, null);
}

function test_growth() {
  const eng = engine();
  const spec = { id: 'g', kind: 'growth', unit: '%', where: { all: [{ eq: ['income_nature', 'rent'] }, NOT_MFO] } };
  let res = eng.compute(spec, { period: '2025' });
  // R = A1,A2,A3,A4,A9 ; I = A1 (80->100), A2 (0->0) ; prev 80, curr 100 -> +25 %
  assert.deepEqual(res.included_ids, ['A1', 'A2']);
  assert.equal(res.prev_sum, 80);
  assert.equal(res.curr_sum, 100);
  assert.equal(res.abs_change, 20);
  assert.ok(close(res.value, 25));
  assert.equal(res.revenue_start, false);
  assert.deepEqual(res.only_curr, ['A4', 'A9']);
  assert.deepEqual(res.only_prev, ['A3']);
  assert.equal(res.status, 'partial');
  assert.equal(res.comparability.status, 'unverified');
  assert.equal(res.prev_period, '2024');
  assert.equal(res.revenue_start, false);
  assert.deepEqual(res.revenue_start_ids, []);
  assert.equal(res.start_change, null);
  // revenue start only: prev 0, curr 50 -> no %, «بدء إيراد» with its riyal change
  const r2 = createEngine({ records: [rec('B1', {}, { 'rev.2025': tv(50), 'rev.2024': tv(0) })], config: config() })
    .compute(spec, { period: '2025' });
  assert.equal(r2.value, null);
  assert.equal(r2.revenue_start, true);
  assert.deepEqual(r2.revenue_start_ids, ['B1']);
  assert.equal(r2.start_change, 50);
  assert.equal(r2.abs_change, null);
  assert.equal(r2.status, 'not_computable');
  assert.equal(r2.reason, lib.REVENUE_START_ONLY_REASON);
  assert.deepEqual(r2.excluded, [{ id: 'B1', reason: 'revenue_start' }]);
  // B9: A 100->100 and B 0->1000 is 0 % for the existing asset plus a separate revenue start (never 1000 %)
  const r7 = createEngine({ records: [rec('A', {}, { 'rev.2025': tv(100), 'rev.2024': tv(100) }),
    rec('B', {}, { 'rev.2025': tv(1000), 'rev.2024': tv(0, 'zero_confirmed') })], config: config() }).compute(spec);
  assert.equal(r7.value, 0);
  assert.deepEqual(r7.included_ids, ['A']);
  assert.equal(r7.status, 'partial');
  assert.deepEqual(r7.revenue_start_ids, ['B']);
  assert.equal(r7.start_change, 1000);
  // B19: prev_sum <= 0 without a revenue start -> not_computable with a reason
  for (const recs of [[rec('N', {}, { 'rev.2025': tv(100), 'rev.2024': tv(-50) })], [rec('Z', {}, { 'rev.2025': tv(0), 'rev.2024': tv(0) })]]) {
    const r8 = createEngine({ records: recs, config: config() }).compute(spec);
    assert.equal(r8.status, 'not_computable');
    assert.equal(r8.value, null);
    assert.equal(r8.reason, lib.PREV_NONPOSITIVE_REASON);
    assert.equal(r8.revenue_start, false);
    assert.equal(r8.n_included, 1);
  }
  // not_comparable
  const cfg = config();
  cfg.periods[0].comparable_with_prev = { status: 'not_comparable', reason: 'partial' };
  const r3 = createEngine({ records: records(), config: cfg }).compute(spec, { period: '2025' });
  assert.equal(r3.status, 'not_comparable');
  assert.equal(r3.value, null);
  assert.equal(r3.n_included, 0);
  assert.equal(r3.reason, 'partial');
  // no previous period -> not_computable (placeholder {REV_PREV} null)
  const r4 = eng.compute(spec, { period: '2024' });
  assert.equal(r4.status, 'not_computable');
  assert.equal(r4.value, null);
  assert.equal(r4.reason, lib.DEFAULT_UNAVAILABLE);
  // explicit prev/curr keys
  const r5 = eng.compute(Object.assign({}, spec, { prev: 'rev.2024', curr: 'rev.2025', where: { eq: ['id', 'A1'] } }));
  assert.ok(close(r5.value, 25));
  // decline: prev 200 -> curr 50 = -75 %
  const r6 = createEngine({ records: [rec('B2', {}, { 'rev.2025': tv(50), 'rev.2024': tv(200) })], config: config() })
    .compute(spec);
  assert.ok(close(r6.value, -75));
}

function test_concentration_tiebreak_and_hhi() {
  const recs = [rec('C3', {}, { 'val.2024': tv(100) }), rec('C1', {}, { 'val.2024': tv(100) }), rec('C2', {}, { 'val.2024': tv(300) }),
    rec('C4', {}, { 'val.2024': tv(null) }), rec('C5', {}, { 'val.2024': tv(0, 'zero_recorded') }), rec('C6', {}, { 'val.2024': tv(100) })];
  const eng = createEngine({ records: recs, config: config() });
  const spec = { id: 'vc', kind: 'concentration', field: '{VAL}', top: 3, unit: '%', where: ASSET };
  let res = eng.compute(spec);
  // K = C1 100, C2 300, C3 100, C6 100 ; total 600 ; ranked C2, C1, C3 (tie by id), C6
  assert.deepEqual(res.top_ids, ['C2', 'C1', 'C3']);
  assert.equal(res.total, 600);
  assert.equal(res.n_included, 4);
  assert.ok(close(res.value, 500 / 600 * 100));
  // hhi = (0.5^2 + 3 * (1/6)^2) * 10000 = 3333.33...
  assert.ok(close(res.hhi, (0.25 + 3 / 36) * 10000, 1e-9));
  assert.equal(res.status, 'partial');                       // C4 unknown
  assert.ok(res.excluded.some((e) => e.id === 'C5' && e.reason === 'nonpositive'));
  assert.ok(res.excluded.some((e) => e.id === 'C4' && e.reason === 'blank'));
  assert.equal(res.top_shares[0].share, 50);
  res = eng.compute(Object.assign({}, spec, { where: { eq: ['id', 'C4'] } }));
  assert.equal(res.status, 'not_computable');
  assert.equal(res.value, null);
  res = eng.compute(Object.assign({}, spec, { where: { in: ['id', ['C1', 'C2', 'C5']] } }));
  assert.equal(res.status, 'complete');                      // known zero does not make it partial
  // top 0 -> Python `0 or 3` = 3 ; top missing -> 3
  res = eng.compute(Object.assign({}, spec, { top: 0 }));
  assert.equal(res.top, 3);
  // ids with mixed case / digits sort by code unit (Python sorted), not locale
  const recs2 = [rec('b', {}, { 'val.2024': tv(5) }), rec('B', {}, { 'val.2024': tv(5) }), rec('a10', {}, { 'val.2024': tv(5) }),
    rec('a9', {}, { 'val.2024': tv(5) })];
  res = createEngine({ records: recs2, config: config() }).compute(Object.assign({}, spec, { top: 4 }));
  assert.deepEqual(res.top_ids, ['B', 'a10', 'a9', 'b']);
}

function test_coverage() {
  const eng = engine();
  const spec = { id: 'cov', kind: 'coverage', fields: ['{REV}', '{VAL}'], unit: '%', where: ASSET };
  let res = eng.compute(spec);
  // R = A1..A6, A9 (7). REV known: A1,A2,A4,A6,A9 = 5 ; VAL known: A1,A2,A3,A5,A6,A9 = 6 -> 5/7*100
  assert.deepEqual(res.per_field, [{ field: 'rev.2025', n_known: 5, n: 7, n_not_applicable: 0 },
    { field: 'val.2024', n_known: 6, n: 7, n_not_applicable: 0 }]);
  assert.equal(res.value, Math.min(5 / 7, 6 / 7) * 100);
  assert.equal(res.status, 'complete');
  assert.equal(res.n_included, 7);
  res = eng.compute(Object.assign({}, spec, { where: { eq: ['id', 'nope'] } }));
  assert.equal(res.status, 'empty');
  assert.equal(res.value, null);
  // B13: «لا ينطبق» leaves that field's denominator
  const recs = records();
  recs[4].fields['rev.2025'] = tv(null, 'not_applicable');      // A5 service asset
  res = createEngine({ records: recs, config: config() }).compute(spec);
  assert.deepEqual(res.per_field[0], { field: 'rev.2025', n_known: 5, n: 6, n_not_applicable: 1 });
  assert.equal(res.value, Math.min(5 / 6, 6 / 7) * 100);
  res = createEngine({ records: recs, config: config() }).compute(Object.assign({}, spec, { fields: ['rev.2025'], where: { eq: ['id', 'A5'] } }));
  assert.equal(res.status, 'not_computable');
  assert.equal(res.value, null);
  assert.equal(res.reason, lib.NO_COVERAGE_FIELD_REASON);
}

function test_trend_same_records() {
  // A18: 2024 known A1 80, A2 0, A3 50 ; 2025 known A1 100, A2 0, A4 300, A9 0 -> the record set changes
  const t = engine().computeTrend();
  assert.equal(t.n_known_varies, true);
  assert.equal(t.same_records.n, 2);
  assert.deepEqual(t.same_records.ids, ['A1', 'A2']);
  assert.deepEqual(t.same_records.points.map((p) => p.value), [80, 100]);
  const t2 = createEngine({ records: [rec('S1', {}, { 'rev.2024': tv(10), 'rev.2025': tv(20) }),
    rec('S2', {}, { 'rev.2024': tv(5), 'rev.2025': tv(0) })], config: config() }).computeTrend();
  assert.equal(t2.n_known_varies, false);
  assert.deepEqual(t2.same_records.points.map((p) => p.value), [15, 20]);
  const t3 = createEngine({ records: [rec('X1', {}, { 'rev.2024': tv(10) }), rec('X2', {}, { 'rev.2025': tv(10) })], config: config() }).computeTrend();
  assert.deepEqual(t3.points.map((p) => p.n_known), [1, 1]);
  assert.equal(t3.n_known_varies, true);                        // same count, different records
  assert.equal(t3.same_records.points[0].value, null);
  const t4 = engine().computeTrend({ location: 'nowhere' });
  assert.equal(t4.status, 'empty');
  assert.equal(t4.empty_cause, 'filters');
  assert.equal(t4.same_records.n, 0);
}

function test_valid_periods() {
  const eng = engine();
  const spec = { id: 'gy', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, unit: '%', where: ASSET,
    valid_periods: ['2025'], invalid_period_reason: 'تقييم 2024 لا يقابل إيراد 2024 هنا' };
  assert.equal(eng.compute(spec, { period: '2025' }).status, 'partial');
  let bad = eng.compute(spec, { period: '2024' });
  assert.equal(bad.status, 'not_comparable');
  assert.equal(bad.value, null);
  assert.equal(bad.reason, 'تقييم 2024 لا يقابل إيراد 2024 هنا');
  assert.equal(bad.n_scope, 7);
  assert.equal(bad.n_included, 0);
  const engNv = createEngine({ records: records(), config: config({ valuation: { field: null } }) });
  bad = engNv.compute(Object.assign({}, spec, { invalid_period_reason: null }), { period: '2024' });
  assert.equal(bad.status, 'not_comparable');
  assert.equal(bad.reason, lib.INVALID_PERIOD_REASON);
  bad = engNv.compute(Object.assign({}, spec, { where: { known: '{VAL}' } }), { period: '2024' });
  assert.equal(bad.n_scope, 8);
  assert.equal(engNv.compute(Object.assign({}, spec, { where: { known: '{VAL}' } }), { period: '2025' }).status, 'not_computable');
  assert.notEqual(eng.compute(Object.assign({}, spec, { valid_periods: null }), { period: '2024' }).status, 'not_comparable');
}

function test_empty_cause_and_empty_as_zero() {
  const eng = engine();
  const zr = { id: 'zr', kind: 'sum', field: '{VAL}', unit: 'SAR', empty_as_zero: true, empty_text: 'لا توجد أصول استثمارية بإيراد صفري مسجل',
    where: { all: [ASSET, { eq: ['use', 'admin'] }, { status: ['{REV}', ['zero_recorded']] }] } };
  let res = eng.compute(zr);
  assert.equal(res.status, 'complete');
  assert.equal(res.value, 0);
  assert.equal(res.empty_cause, 'scope');
  assert.equal(res.note, zr.empty_text);
  assert.equal(res.reason, null);
  res = eng.compute(Object.assign({}, zr, { empty_as_zero: false }));
  assert.equal(res.status, 'empty');
  assert.equal(res.value, null);
  assert.equal(res.reason, zr.empty_text);
  const op = { id: 'op', kind: 'sum', field: '{REV}', unit: 'SAR', empty_as_zero: true, where: { eq: ['income_nature', 'operating'] } };
  res = eng.compute(op, { location: 'شمال' });
  assert.equal(res.status, 'empty');
  assert.equal(res.empty_cause, 'filters');
  assert.equal(res.reason, lib.EMPTY_FILTERS_REASON);
  res = eng.compute({ id: 'r', kind: 'ratio', num: '{REV}', den: '{VAL}', unit: '%', empty_as_zero: true, where: { eq: ['record_type', 'unit'] } });
  assert.equal(res.status, 'empty');
  assert.equal(res.reason, lib.EMPTY_SCOPE_REASON);
  const bd = { id: 'b', label: 'b', field: '{VAL}', by: 'use', where: { eq: ['record_type', 'unit'] }, empty_text: 'لا وحدات' };
  res = eng.computeBreakdown(bd);
  assert.equal(res.status, 'empty');
  assert.equal(res.empty_cause, 'scope');
  assert.equal(res.reason, 'لا وحدات');
  res = eng.computeBreakdown(Object.assign({}, bd, { where: ASSET }), { location: 'nowhere' });
  assert.equal(res.empty_cause, 'filters');
}

function test_unresolved_placeholder() {
  const eng = engine(null, { valuation: { field: null, year: null } });
  const spec = { id: 'rv', kind: 'sum', field: '{VAL}', unit: 'SAR', where: ASSET, unavailable_reason: 'لا يوجد تقييم' };
  let res = eng.compute(spec);
  assert.equal(res.status, 'not_computable');
  assert.equal(res.reason, 'لا يوجد تقييم');
  assert.equal(res.value, null);
  assert.equal(res.n_scope, 7);                 // |R| (base_where AND where)
  assert.equal(res.n_included, 0);
  assert.deepEqual(res.excluded, []);
  // placeholder inside where: n_scope falls back to base_where AND filters (8 in-scope records)
  res = eng.compute(Object.assign({}, spec, { field: 'rev.2025', where: { known: '{VAL}' } }));
  assert.equal(res.status, 'not_computable');
  assert.equal(res.n_scope, 8);
  res = eng.compute(Object.assign({}, spec, { unavailable_reason: null }));
  assert.equal(res.reason, lib.DEFAULT_UNAVAILABLE);
  res = eng.compute(Object.assign({}, spec, { unavailable_reason: '' }));
  assert.equal(res.reason, lib.DEFAULT_UNAVAILABLE);
  // unknown kind throws
  assert.throws(() => eng.compute({ id: 'k', kind: 'median', field: 'x' }));
}

function test_breakdown_list_dim_and_none() {
  const recs = records().concat([rec('A10', { own: [], atype: null }, { 'rev.2025': tv(20), 'val.2024': tv(10) })]);
  const eng = createEngine({ records: recs, config: config() });
  const spec = { id: 'v_own', label: 'x', field: '{VAL}', by: 'ownership', where: ASSET, unit: 'SAR' };
  const res = eng.computeBreakdown(spec);
  assert.equal(res.overlap, true);
  assert.equal(res.note, lib.OVERLAP_NOTE);
  const owned = res.rows.find((r) => r.key === 'owned');
  // owned: A1 1000, A2 500, A3 2000, A5 700, A9 -1 known (4199) ; A4 unknown
  assert.equal(owned.value, 4199);
  assert.equal(owned.n_known, 5);
  assert.equal(owned.n_unknown, 1);
  const none = res.rows.find((r) => r.key === lib.NONE_KEY);
  assert.equal(none.value, 10);
  assert.equal(res.total, 4199 + 999 + 10);
  assert.ok(close(owned.share, 4199 / res.total));
  assert.equal(res.rows[0].key, 'owned');
  const res2 = createEngine({ records: recs, config: config() }).computeBreakdown(Object.assign({}, spec, { by: 'asset_type', id: 'v_type' }));
  assert.ok(res2.rows.some((r) => r.key === lib.NONE_KEY));
  const land = res2.rows.find((r) => r.key === 'أرض');
  assert.equal(land.value, null);               // only unknown values -> null, sorted last
  assert.equal(land.n_unknown, 1);
  assert.equal(res2.rows[res2.rows.length - 1].key, 'أرض');
  assert.equal(res2.status, 'partial');
  assert.equal(res2.overlap, false);
  // empty
  const res3 = eng.computeBreakdown(spec, { location: 'nowhere' });
  assert.equal(res3.status, 'empty');
  // unresolved placeholder
  const res4 = createEngine({ records: recs, config: config({ valuation: { field: null } }) }).computeBreakdown(spec);
  assert.equal(res4.status, 'not_computable');
  assert.equal(res4.unresolved, '{VAL}');
}

function test_trend() {
  const eng = engine();
  const t = eng.computeTrend();
  // R = rent, not mfo, in scope: A1,A2,A3,A4,A9 ; 2024: 80+0+50 = 130 (3 known) ; 2025: 100+0+300+0 = 400 (4 known)
  assert.deepEqual(t.points[0], { period: '2024', field: 'rev.2024', value: 130, n_known: 3, n: 5 });
  assert.equal(t.points[1].value, 400);
  assert.equal(t.points[1].n_known, 4);
  assert.equal(t.status, 'partial');
  assert.equal(t.not_comparable, false);
  assert.equal(t.unit, 'SAR');
  const cfg = config();
  cfg.trend.comparable = { status: 'not_comparable', reason: 'partial year' };
  const t2 = createEngine({ records: records(), config: cfg }).computeTrend();
  assert.equal(t2.not_comparable, true);
  assert.equal(t2.points[0].value, 130);        // values kept, flagged
  assert.equal(eng.computeTrend({ location: 'nowhere' }).status, 'empty');
  const t4 = createEngine({ records: records(), config: config({ trend: { enabled: false } }) }).computeTrend();
  assert.equal(t4.status, 'not_computable');
}

function test_filters_and_compute_all() {
  const eng = engine();
  assert.deepEqual(eng.normalizeFilters({ use: '', location: null, asset_type: 'عمارة', x: undefined }),
    { period: '2025', asset_type: 'عمارة' });
  assert.deepEqual(eng.filterRecords({ location: 'شمال' }).map((r) => r.id), ['A3', 'A5']);
  // ownership is a list dim -> has
  assert.deepEqual(eng.filterRecords({ ownership: 'usufruct' }).map((r) => r.id), ['A7']);
  // base_where removes the isolated A8
  assert.deepEqual(eng.filterRecords({}).map((r) => r.id), ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A9']);
  const all = eng.computeAll({ period: '2024', use: 'investment' });
  assert.deepEqual(all.filters, { period: '2024', use: 'investment' });
  assert.ok(!all.records.includes('A5') && !all.records.includes('A8'));
  assert.deepEqual(all.record_ids, all.records);
  assert.throws(() => eng.normalizeFilters({ period: '1999' }));
  assert.throws(() => eng.normalizeFilters({ period: 2025 }));   // strict: number is not the string id
  // a boolean dim filter uses strict equality
  assert.deepEqual(eng.filterRecords({ managed_for_others: true }).map((r) => r.id), ['A6']);
}

function test_text_by_period() {
  const eng = engine();
  const spec = { id: 'pr', kind: 'sum', field: '{REV}', unit: 'SAR', label: 'إيراد 2025',
    text_by_period: { 2024: { label: 'إيراد 2024' } } };
  assert.equal(eng.compute(spec, { period: '2025' }).label, 'إيراد 2025');
  assert.equal(eng.compute(spec, { period: '2024' }).label, 'إيراد 2024');
  const r = eng.compute(spec);
  assert.equal(r.card, false);
  assert.equal(r.definition, null);             // missing text keys are null, not undefined
}

function test_standard_metric_set() {
  // The standard metric ids/kinds of compute_metrics.py --init-config, on the fixture.
  const W_ASSET = ASSET, W_NOT_MANAGED = NOT_MFO, W_RENT = { eq: ['income_nature', 'rent'] };
  const W_INVEST = { eq: ['use', 'investment'] }, W_NOT_OPERATING = { not: { eq: ['income_nature', 'operating'] } };
  const metrics = [
    { id: 'asset_count', kind: 'count', where: W_ASSET, unit: 'asset', card: true, card_group: 'counts', card_order: 1 },
    { id: 'operating_records_count', kind: 'count', where: { eq: ['record_type', 'operating_activity'] }, unit: 'count', card: true, card_group: 'counts', card_order: 2 },
    { id: 'recorded_value', kind: 'sum', field: '{VAL}', where: { all: [W_ASSET, W_NOT_MANAGED] }, unit: 'SAR', card: true },
    { id: 'property_revenue', kind: 'sum', field: '{REV}', where: { all: [W_RENT, W_NOT_MANAGED] }, unit: 'SAR', card: true },
    { id: 'activity_revenue', kind: 'sum', field: '{REV}', where: { eq: ['income_nature', 'operating'] }, unit: 'SAR', card: true },
    { id: 'gross_yield_investment', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, where: { all: [W_INVEST, W_ASSET, W_NOT_OPERATING, W_NOT_MANAGED] }, unit: '%', card: true },
    { id: 'gross_yield_income_producing', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, where: { all: [W_INVEST, W_ASSET, W_NOT_OPERATING, W_NOT_MANAGED, { gt: ['{REV}', 0] }] }, unit: '%' },
    { id: 'noi', kind: 'diff', a: '{REV}', b: '{EXP}', where: { all: [W_RENT, W_NOT_MANAGED] }, unit: 'SAR', card: true },
    { id: 'zero_rev_investment_value', kind: 'sum', field: '{VAL}', where: { all: [W_INVEST, W_ASSET, W_NOT_MANAGED, { status: ['{REV}', ['zero_recorded', 'zero_confirmed']] }] }, unit: 'SAR', card: true },
    { id: 'data_coverage', kind: 'coverage', fields: ['{REV}', '{VAL}'], where: W_ASSET, unit: '%', card: true },
    { id: 'revenue_growth', kind: 'growth', prev: '{REV_PREV}', curr: '{REV}', where: { all: [W_RENT, W_NOT_MANAGED] }, unit: '%' },
    { id: 'value_concentration_top3', kind: 'concentration', field: '{VAL}', top: 3, where: { all: [W_ASSET, W_NOT_MANAGED] }, unit: '%' },
    { id: 'revenue_concentration_top3', kind: 'concentration', field: '{REV}', top: 3, where: W_NOT_MANAGED, unit: '%' }
  ];
  const eng = engine(null, { metrics });
  const m = eng.computeAll({ period: '2025' }).metrics;
  assert.equal(m.asset_count.value, 7);
  assert.equal(m.operating_records_count.value, 1);              // A7
  assert.equal(m.recorded_value.value, 4199);
  assert.equal(m.property_revenue.value, 400);
  assert.equal(m.activity_revenue.value, 400);                   // A7
  assert.equal(m.activity_revenue.status, 'complete');
  assert.ok(close(m.gross_yield_investment.value, 100 / 1500 * 100));
  // income producing: gt REV 0 -> A1 (100/1000), A4 (den unknown) -> 10 %, partial
  assert.ok(close(m.gross_yield_income_producing.value, 10));
  assert.equal(m.gross_yield_income_producing.status, 'partial');
  assert.equal(m.noi.value, 85);
  assert.equal(m.zero_rev_investment_value.value, 499);
  assert.ok(close(m.data_coverage.value, 5 / 7 * 100));
  assert.ok(close(m.revenue_growth.value, 25));
  // value concentration: K = A1 1000, A2 500, A3 2000, A5 700 (A9 -1 nonpositive, A4 unknown); top3 = 3700 / 4200
  assert.deepEqual(m.value_concentration_top3.top_ids, ['A3', 'A1', 'A5']);
  assert.ok(close(m.value_concentration_top3.value, 3700 / 4200 * 100));
  // revenue concentration over not-managed in scope: A1 100, A4 300, A7 400 positive; total 800 -> top3 = 100 %
  assert.ok(close(m.revenue_concentration_top3.value, 100));
  assert.equal(m.revenue_concentration_top3.status, 'partial');  // A3 blank, A5 absent
  // 2024: no expense field -> noi not_computable ; growth not_computable
  const m24 = eng.computeAll({ period: '2024' }).metrics;
  assert.equal(m24.noi.status, 'not_computable');
  assert.equal(m24.revenue_growth.status, 'not_computable');
  assert.equal(m24.property_revenue.value, 130);
  // filters narrow everything consistently
  const mn = eng.computeAll({ period: '2025', location: 'شمال' }).metrics;
  assert.equal(mn.asset_count.value, 2);                          // A3, A5
  assert.equal(mn.activity_revenue.status, 'empty');
  assert.equal(mn.activity_revenue.value, null);
  assert.equal(mn.asset_count.status, 'complete');
}

function test_determinism_independent_of_input_order() {
  const specs = [
    { id: 's', kind: 'sum', field: '{VAL}', unit: 'SAR', where: ASSET },
    { id: 'r', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, unit: '%', where: ASSET },
    { id: 'c', kind: 'concentration', field: '{REV}', top: 2, unit: '%' },
    { id: 'g', kind: 'growth', unit: '%' }
  ];
  const base = engine(null, { metrics: specs }).computeAll();
  let seed = 7;
  const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  for (let k = 0; k < 5; k++) {
    const recs = records();
    for (let i = recs.length - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [recs[i], recs[j]] = [recs[j], recs[i]]; }
    const other = createEngine({ records: recs, config: config({ metrics: specs }) }).computeAll();
    for (const id of Object.keys(base.metrics)) {
      assert.equal(other.metrics[id].value, base.metrics[id].value);
      assert.deepEqual(other.metrics[id].included_ids, base.metrics[id].included_ids);
      assert.deepEqual(other.metrics[id].excluded, base.metrics[id].excluded);
    }
    assert.deepEqual(other.records, base.records);
  }
}

function test_does_not_mutate_input() {
  const data = { records: records(), config: config({ metrics: [{ id: 's', kind: 'sum', field: '{VAL}', where: ASSET }] }) };
  const before = JSON.stringify(data);
  const eng = createEngine(data);
  eng.computeAll({ period: '2024', use: 'investment' });
  eng.computeAll();
  assert.equal(JSON.stringify(data), before);
}

function test_loading_modes() {
  const src = fs.readFileSync(ENGINE_PATH, 'utf8');
  assert.ok(!/\beval\s*\(|new Function\s*\(|fetch\s*\(|XMLHttpRequest|import\s*\(/.test(src), 'no eval / network in the engine');
  // browser-like: window.createEngine
  const win = {};
  const ctxB = vm.createContext({ window: win });
  vm.runInContext(src, ctxB);
  assert.equal(typeof win.createEngine, 'function');
  assert.equal(win.createEngine({ records: records(), config: config() }).filterRecords({}).length, 8);
  // bare script without CommonJS / window: globalThis.createEngine
  const ctxG = vm.createContext({});
  vm.runInContext(src, ctxG);
  assert.equal(typeof ctxG.createEngine, 'function');
  // inlined inside a template script that later builds window.APA (as build_dashboard.py does)
  const page = '/*engine*/' + src + '\nconst DATA = ' + JSON.stringify({ records: records(), config: config() }).replace(/<\//g, '<\\/') +
    ';\nwindow.APA = Object.assign(createEngine(DATA), { data: DATA });';
  const win2 = {};
  const ctxP = vm.createContext({ window: win2 });
  // in a browser, window properties are globals; emulate by making the context global the window
  vm.runInContext('var window = this;', ctxP);
  vm.runInContext(page, ctxP);
  assert.equal(ctxP.APA.computeAll().records.length, 8);
}

// ---------------------------------------------------------------------------
// Cross-check against the Python engine on the same fixture (skipped if unavailable)
// ---------------------------------------------------------------------------
function cross_records() {
  return records().concat([
    rec('A10', { own: [], atype: null, loc: 'شمال' }, { 'rev.2025': tv(20), 'val.2024': tv(10), 'exp.2025': tv(1, 'zero_confirmed') }),
    rec('A11', { own: ['owned', 'partial'], use: 'unknown', rtype: 'land_parcel' },
      { 'rev.2025': { v: 12.5, s: 'derived' }, 'rev.2024': tv(10), 'val.2024': tv(250.75), 'exp.2025': tv(2.25) }),
    rec('A12', { use: 'admin', inc: 'none' }, { 'rev.2025': { v: null, s: 'text', raw: 'غير مقيد' }, 'val.2024': tv(0, 'zero_recorded') }),
    rec('ب1', { atype: 'محل', loc: 'وسط' }, { 'rev.2025': tv(1 / 3), 'rev.2024': tv(0.1), 'val.2024': tv(0.7) }),
    // revenue start (prev recorded zero, curr > 0): outside the growth %, in revenue_start_ids / start_change
    rec('A13', { loc: 'وسط' }, { 'rev.2025': tv(40), 'rev.2024': tv(0), 'val.2024': tv(400) }),
    rec('A16', { loc: 'شمال' }, { 'rev.2025': tv(7.5), 'rev.2024': tv(0, 'zero_confirmed') }),
    // «لا ينطبق» revenue on an admin asset: leaves the coverage denominator
    rec('A14', { use: 'admin', inc: 'none' }, { 'rev.2025': { v: null, s: 'not_applicable', raw: 'لا ينطبق' }, 'val.2024': tv(300) }),
    // negative previous revenue: growth over it alone is not computable
    rec('A15', { use: 'unknown', loc: 'وسط' }, { 'rev.2025': tv(5), 'rev.2024': tv(-10) }),
    // a unit row (never matched by the revenue metrics of compute_metrics; here only by the unit metrics)
    rec('A17', { rtype: 'unit' }, { 'rev.2025': tv(3), 'val.2024': tv(30) })
  ]);
}

function cross_config() {
  const W_RENT = { eq: ['income_nature', 'rent'] };
  return config({
    metrics: [
      { id: 'cnt', kind: 'count', where: ASSET, card: true, card_group: 'counts', card_order: 1 },
      { id: 'cnt_op', kind: 'count', where: { eq: ['record_type', 'operating_activity'] } },
      { id: 'val', kind: 'sum', field: '{VAL}', where: { all: [ASSET, NOT_MFO] }, unit: 'SAR' },
      { id: 'rev', kind: 'sum', field: '{REV}', where: { all: [W_RENT, NOT_MFO] }, unit: 'SAR', text_by_period: { 2024: { label: 'L24' } } },
      { id: 'yield', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, where: { all: [ASSET, NOT_MFO] } },
      { id: 'yield_pos', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, where: { all: [ASSET, { gt: ['{REV}', 0] }] } },
      { id: 'noi', kind: 'diff', a: '{REV}', b: '{EXP}', where: { all: [W_RENT, NOT_MFO] }, unavailable_reason: 'no exp' },
      { id: 'zero_val', kind: 'sum', field: '{VAL}', where: { all: [ASSET, { status: ['{REV}', ['zero_recorded', 'zero_confirmed']] }] } },
      { id: 'cov', kind: 'coverage', fields: ['{REV}', '{VAL}', '{EXP}'], where: ASSET },
      { id: 'growth', kind: 'growth', where: { all: [W_RENT, NOT_MFO] } },
      { id: 'conc_val', kind: 'concentration', field: '{VAL}', top: 3, where: ASSET },
      { id: 'conc_rev', kind: 'concentration', field: '{REV}', top: 2 },
      { id: 'has_owned', kind: 'count', where: { all: [{ has: ['ownership', 'owned'] }, { not: { in: ['use', ['admin', 'service']] } }] } },
      { id: 'known_exp', kind: 'sum', field: 'rev.2025', where: { any: [{ known: '{EXP}' }, { eq: ['asset_type', null] }] } },
      // C2 valid_periods (2024 -> not_comparable), with and without a reason, and over an unresolved placeholder
      { id: 'yield_vp', kind: 'ratio', num: '{REV}', den: '{VAL}', scale: 100, where: ASSET, valid_periods: ['2025'],
        invalid_period_reason: 'تقييم 2024 لا يقابل إيراد هذه الفترة' },
      { id: 'noi_vp', kind: 'diff', a: '{REV}', b: '{EXP}', valid_periods: ['2025'] },
      // C3 empty cause / empty_as_zero / empty_text
      { id: 'units_val', kind: 'sum', field: '{VAL}', where: { eq: ['record_type', 'contract'] }, empty_as_zero: true, empty_text: 'لا عقود' },
      { id: 'units_val_nz', kind: 'sum', field: '{VAL}', where: { eq: ['record_type', 'contract'] }, empty_text: 'لا عقود' },
      { id: 'cnt_contract', kind: 'count', where: { eq: ['record_type', 'contract'] } },
      { id: 'cnt_contract_nz', kind: 'count', where: { eq: ['record_type', 'contract'] }, empty_as_zero: false },
      { id: 'unit_rev', kind: 'sum', field: '{REV}', where: { eq: ['record_type', 'unit'] }, empty_as_zero: true },
      { id: 'ratio_empty', kind: 'ratio', num: '{REV}', den: '{VAL}', where: { eq: ['record_type', 'contract'] }, empty_as_zero: true },
      { id: 'zero_rev', kind: 'sum', field: '{VAL}', empty_as_zero: true, empty_text: 'لا أصول بإيراد صفري',
        where: { all: [ASSET, { eq: ['use', 'investment'] }, { status: ['{REV}', ['zero_recorded', 'zero_confirmed']] }] } },
      // B9/B19 growth edge cases
      { id: 'growth_all', kind: 'growth', where: ASSET },
      { id: 'growth_neg', kind: 'growth', where: { eq: ['id', 'A15'] } },
      { id: 'growth_start', kind: 'growth', where: { in: ['id', ['A13', 'A16']] } },
      // B13 coverage with a not-applicable field value
      { id: 'cov_na', kind: 'coverage', fields: ['{REV}', '{VAL}'], where: { all: [ASSET, { in: ['use', ['admin', 'investment']] }] } },
      { id: 'cov_all_na', kind: 'coverage', fields: ['{REV}'], where: { eq: ['id', 'A14'] } }
    ],
    breakdowns: [
      { id: 'b_type', field: '{VAL}', by: 'asset_type', where: ASSET },
      { id: 'b_own', field: '{REV}', by: 'ownership' },
      { id: 'b_mfo', field: '{VAL}', by: 'managed_for_others' },
      { id: 'b_exp', field: '{EXP}', by: 'use' },
      { id: 'b_contract', field: '{VAL}', by: 'use', where: { eq: ['record_type', 'contract'] }, empty_text: 'لا عقود' }
    ]
  });
}

function test_cross_check_with_python_engine() {
  const pyEngine = path.join(SKILL, 'scripts', 'metric_engine.py');
  if (!fs.existsSync(pyEngine)) return 'SKIP (scripts/metric_engine.py not found)';
  const filtersList = [{}, { period: '2024' }, { use: 'investment' }, { location: 'شمال' }, { asset_type: 'محل' },
    { ownership: 'owned' }, { ownership: 'partial', period: '2024' }, { use: 'admin' }, { location: 'nowhere' },
    { managed_for_others: true }, { use: '', location: null }];
  const payload = JSON.stringify({ records: cross_records(), config: cross_config(), filters: filtersList });
  const code = [
    'import json, sys',
    'sys.path.insert(0, sys.argv[1])',
    'import metric_engine as me',
    'p = json.loads(sys.stdin.read())',
    'eng = me.Engine(p["config"], p["records"])',
    'print(json.dumps([eng.compute_all(f) for f in p["filters"]], ensure_ascii=False))'
  ].join('\n');
  const py = spawnSync('python3', ['-I', '-B', '-c', code, path.join(SKILL, 'scripts')], { input: payload, encoding: 'utf8', maxBuffer: 64 << 20 });
  if (py.error) return 'SKIP (python3 not available)';
  assert.equal(py.status, 0, 'python engine failed: ' + py.stderr);
  const pyRes = JSON.parse(py.stdout);
  const eng = createEngine({ records: cross_records(), config: cross_config() });
  let compared = 0;
  const diffs = [];
  const cmp = (a, b, p) => {
    if (typeof a === 'number' && typeof b === 'number') {
      compared++;
      // bit-identical: same accumulation order and the same operations (e.g. HHI uses s*s in both engines)
      if (a !== b) diffs.push(p + ': py=' + a + ' js=' + b);
      return;
    }
    if (a === null || b === null || typeof a !== 'object' || typeof b !== 'object') {
      compared++;
      if (a !== b) diffs.push(p + ': py=' + JSON.stringify(a) + ' js=' + JSON.stringify(b));
      return;
    }
    if (Array.isArray(a) !== Array.isArray(b)) { diffs.push(p + ': array vs object'); return; }
    const keys = new Set(Object.keys(a).concat(Object.keys(b)));
    for (const k of keys) cmp(a[k], b[k], p + '.' + k);
  };
  filtersList.forEach((f, i) => {
    const js = eng.computeAll(f);
    delete js.record_ids;                       // JS-only alias of `records`
    cmp(pyRes[i], js, JSON.stringify(f));
  });
  assert.deepEqual(diffs, [], diffs.slice(0, 10).join('\n'));
  return 'compared ' + compared + ' values over ' + filtersList.length + ' filter sets';
}

// ---------------------------------------------------------------------------
const TESTS = {
  test_where_ops, test_placeholders_in_where, test_count, test_sum_statuses, test_unknown_statuses_never_summed,
  test_ratio_zero_included_and_exclusion_reasons, test_diff_and_missing_expense, test_growth,
  test_concentration_tiebreak_and_hhi, test_coverage, test_trend_same_records, test_valid_periods,
  test_empty_cause_and_empty_as_zero, test_unresolved_placeholder, test_breakdown_list_dim_and_none,
  test_trend, test_filters_and_compute_all, test_text_by_period, test_standard_metric_set,
  test_determinism_independent_of_input_order, test_does_not_mutate_input, test_loading_modes,
  test_cross_check_with_python_engine
};

let failed = 0;
for (const [name, fn] of Object.entries(TESTS)) {
  try {
    const note = fn();
    console.log('PASS ' + name + (note ? ' — ' + note : ''));
  } catch (e) {
    failed++;
    console.log('FAIL ' + name + ': ' + (e && e.message ? e.message : e));
  }
}
const total = Object.keys(TESTS).length;
console.log((total - failed) + '/' + total + ' passed');
process.exit(failed ? 1 : 0);
