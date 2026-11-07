#!/usr/bin/env node
/**
 * check_dashboard.mjs — headless-browser test of a built dashboard.html (data-contract.md §0, §8).
 *
 * Usage:
 *   node scripts/check_dashboard.mjs --html RUN/dashboard.html --out RUN [--timeout 30000] [--allow-draft]
 *   node scripts/check_dashboard.mjs --self-test     (number parsing and shared PII patterns, no browser)
 * Writes RUN/08_dashboard_check.json and RUN/screens/{desktop.png, desktop_01.png …, details_open.png and its
 * readable slices details_open_01.png … (≤ 1400 px each), mobile.png,
 * mobile_full.png, mobile_360.png, mobile_detail.png, print.pdf}.
 * Exit code: 0 = every check passed, 1 = a check failed (or the browser could not run), 2 = bad arguments.
 *
 * Checks (DOM contract shared with assets/dashboard_template.html):
 *   load, draft (a --draft build fails unless --allow-draft), html_lang_dir (attributes and computed direction),
 *   api, structure, details_sections, parity (every parity case of 05_metrics.json recomputed in the page by
 *   APA.computeAll), cards (data-value/status AND the displayed number; ratio cards show their scope, basis and
 *   valuation date), charts (every [data-role=chart-number] against computeAll and its own text), table_cells
 *   (value/revenue/yield cells against the records), filters, period_filter (cards, charts and table cells per
 *   state), table_search, table_sort (every sortable column; numeric columns must be ordered), table_detail,
 *   buttons, no_raw_tokens, privacy (the shared patterns of assets/pii_patterns.json: block level fails, review
 *   level is reported in `warnings`), desktop_screenshot, mobile_overflow (360 and 390 px, default view and with the
 *   details and an asset detail open), growth_no_pct (an unconfirmed growth never shows a %), no_raw_keys (no engine
 *   keys in visible text outside <code>), bidi_refs («!» between an Arabic sheet name and its cell), print (executive
 *   sections visible on paper, yield cards keep scope/basis/valuation date, at most 8 A4 pages), offline, console_errors.
 * Playwright is loaded from the local project, NODE_PATH or the global npm root.
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';

const PARITY_REL_TOL = 1e-9;
const MAX_FILTER_VALUES = 3;
const MAX_BUTTONS = 150;
const MAX_SORT_HEADERS = 30;
const MAX_TABLE_ROWS = 80;
const MAX_CROPS = 14;
const MIN_TAIL = 200;          // px: a shorter remainder joins the previous screenshot crop
const MAX_PRINT_PAGES = 8;     // dashboard-spec §8: the executive print fits in 8 A4 pages
const DETAIL_SECTIONS = ['calculations', 'issues', 'exclusions', 'assumptions', 'sources', 'workbook', 'questions',
  'scenarios', 'versions', 'methodology', 'records'];
// sections of the executive view that must be on paper (brief §11: the executive view prints)
const PRINT_SECTIONS = ['sec-summary', 'sec-kpis', 'sec-priorities', 'sec-charts', 'sec-plan'];
const MOBILE_WIDTHS = [360, 390];
const MAX_CARDS = 8;
const SKILL_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const PII_PATTERNS_PATH = path.join(SKILL_DIR, 'assets', 'pii_patterns.json');

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------
function usage(msg) {
  if (msg) console.error('خطأ: ' + msg);
  console.error('الاستخدام: node scripts/check_dashboard.mjs --html RUN/dashboard.html --out RUN [--timeout 30000] [--allow-draft]');
  process.exit(2);
}

function parseArgs(argv) {
  const out = { timeout: 30000, allowDraft: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const eq = a.indexOf('=');
    const key = eq > 0 ? a.slice(0, eq) : a;
    const val = () => (eq > 0 ? a.slice(eq + 1) : argv[++i]);
    if (key === '--html') out.html = val();
    else if (key === '--out') out.out = val();
    else if (key === '--timeout') out.timeout = Number(val());
    else if (key === '--allow-draft') out.allowDraft = true;
    else if (key === '--self-test') out.selfTest = true;
    else if (key === '-h' || key === '--help') usage();
    else usage('وسيط غير معروف: ' + a);
  }
  if (out.selfTest) return out;
  if (!out.html) usage('--html مطلوب');
  if (!out.out) out.out = path.dirname(path.resolve(out.html));
  if (!Number.isFinite(out.timeout) || out.timeout <= 0) usage('--timeout يجب أن يكون رقمًا موجبًا بالملّي ثانية');
  return out;
}

async function loadPlaywright() {
  const names = ['playwright', 'playwright-core', '@playwright/test'];
  for (const n of names) {
    try {
      const m = await import(n);
      const lib = m.chromium ? m : m.default;
      if (lib && lib.chromium) return lib;
    } catch (e) { /* try next */ }
  }
  const roots = [import.meta.url];
  for (const p of (process.env.NODE_PATH || '').split(path.delimiter).filter(Boolean)) roots.push(path.join(p, 'noop.js'));
  try {
    const g = execSync('npm root -g', { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'], timeout: 15000 }).trim();
    if (g) roots.push(path.join(g, 'noop.js'));
  } catch (e) { /* npm not available */ }
  for (const root of roots) {
    const req = createRequire(root);
    for (const n of names) {
      try {
        const lib = req(n);
        if (lib && lib.chromium) return lib;
      } catch (e) { /* try next */ }
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Small helpers (Node side)
// ---------------------------------------------------------------------------
function numClose(a, b) {
  if (a === null || a === undefined || b === null || b === undefined) return (a ?? null) === (b ?? null);
  if (typeof a !== 'number' || typeof b !== 'number') return a === b;
  if (a === b) return true;
  return Math.abs(a - b) <= PARITY_REL_TOL * Math.max(Math.abs(a), Math.abs(b)) || Math.abs(a - b) <= 1e-12;
}

function cleanFilters(f) {
  const out = {};
  for (const k of Object.keys(f || {}).sort()) {
    const v = f[k];
    if (v === null || v === undefined || v === '') continue;
    out[k] = v;
  }
  return out;
}

const sameJSON = (a, b) => JSON.stringify(a) === JSON.stringify(b);
/** Arabic counted noun: 1 and 2 by their own forms, 3–10 plural, 11+ singular accusative. */
const nAr = (n, one, two, few, many) => (n === 1 ? one : n === 2 ? two : `${n} ${n >= 3 && n <= 10 ? few : many}`);
const short = (x, n = 300) => { const s = typeof x === 'string' ? x : JSON.stringify(x); return s && s.length > n ? s.slice(0, n) + '…' : s; };

function toAsciiDigits(s) {
  return String(s).replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x0660)).replace(/[۰-۹]/g, (d) => String(d.charCodeAt(0) - 0x06F0));
}

/** True when a card text renders a zero (e.g. "0", "0 ريال", "٠٫٠٪"). */
function looksZero(text) {
  const t = toAsciiDigits(text).replace(/ريال(\s*سعودي)?|ر\.?\s*س\.?|SAR|﷼|٪|%/gi, ' ');
  const nums = t.match(/\d[\d.,٫٬]*/g) || [];
  return nums.some((n) => Number(n.replace(/[,٬]/g, '').replace('٫', '.')) === 0);
}

/** A table cell that shows "no value" (dash, empty, «غير متاح» ...). */
function isMissingText(text) {
  const t = String(text).trim();
  return !/[\p{L}\p{N}]/u.test(t) || /^(غير متاح|غير متوفر|لا يوجد|غير محدد|غير معروف|غير مسجل|لا ينطبق)$/.test(t);
}

function hasArabicWords(text) {
  const t = String(text).replace(/ريال(\s*سعودي)?|ر\.?\s*س\.?/g, ' ');
  return /[ء-ي]{2,}/.test(t);
}

/**
 * The first number shown in a text, as the reader sees it: Arabic-Indic digits, thousands separators, the minus
 * sign (ASCII or U+2212) and bidi control characters are handled. Returns {value, decimals} or null.
 */
function shownNumber(text) {
  const t = toAsciiDigits(String(text)).replace(/[⁦-⁩‎‏؜]/g, '').replace(/−/g, '-')
    .replace(/٫/g, '.').replace(/(\d)[,٬](?=\d{3}(?:\D|$))/g, '$1');
  const m = /(-?)\s?(\d+(?:\.\d+)?)/.exec(t);
  if (!m) return null;
  const dec = m[2].includes('.') ? m[2].split('.')[1].length : 0;
  return { value: Number((m[1] ? '-' : '') + m[2]), decimals: dec };
}

/** shownNumber() only when the text starts with the number (a status such as «غير متاح «6 محلات»» is not a number). */
function leadingNumber(text) {
  const t = toAsciiDigits(String(text)).replace(/[⁦-⁩‎‏؜]/g, '').replace(/−/g, '-').trim();
  return /^-?\s?\d/.test(t) ? shownNumber(t) : null;
}

/** True when a displayed number is the rounded form of `expected` (rounding to the decimals it shows). */
function shownMatches(text, expected) {
  const s = shownNumber(text);
  if (!s || typeof expected !== 'number') return false;
  const tol = 0.5 * Math.pow(10, -s.decimals) + 1e-9 * Math.max(1, Math.abs(expected));
  return Math.abs(s.value - expected) <= tol;
}

/** Shared personal-data patterns (assets/pii_patterns.json), compiled for JS; values are tested on ASCII digits. */
function loadPiiPatterns() {
  const raw = JSON.parse(fs.readFileSync(PII_PATTERNS_PATH, 'utf8'));
  const flags = (f) => (String(f || '').includes('i') ? 'i' : '');
  const values = (raw.values || []).map((v) => ({ kind: v.kind, level: v.level || 'block', re: new RegExp(v.pattern, flags(v.flags)) }));
  const hits = (text, level) => {
    const t = toAsciiDigits(text);
    return values.filter((v) => v.level === level && v.re.test(t)).map((v) => v.kind);
  };
  // the same self-test as apa_common._self_test_pii: a pattern that behaves differently in JS must not pass silently
  const selfTest = [];
  const tests = raw.tests || {};
  for (const x of tests.must_match || []) if (!values.some((v) => v.re.test(toAsciiDigits(x)))) selfTest.push('لا يطابق: ' + x.replace(/\d/g, '#'));
  for (const x of tests.must_not_match || []) if (values.some((v) => v.re.test(toAsciiDigits(x)))) selfTest.push('يطابق خطأً: ' + x);
  const blockHit = (x) => values.some((v) => v.level === 'block' && v.re.test(toAsciiDigits(x)));
  for (const x of tests.block_must_match || []) if (!blockHit(x)) selfTest.push('لا يطابق (حجب): ' + x.replace(/\d/g, '#'));
  for (const x of tests.block_must_not_match || []) if (blockHit(x)) selfTest.push('يطابق خطأً (حجب): ' + x);
  return { values, hits, selfTest, version: raw.version };
}

/** A raw cell text that is only the printed form of its numeric value (not scanned: a 9-digit amount may start with 5). */
function rawIsPlainNumber(raw, v) {
  const t = toAsciiDigits(raw).trim();
  if (/^(?:\+|00|0\d)/.test(t)) return false;   // a phone written in a numeric cell keeps its leading 0 or +
  const n = Number(t.replace(/[,\s٬]/g, ''));
  return Number.isFinite(n) && n === v;
}

// ---------------------------------------------------------------------------
// In-page helpers (serialised into the page)
// ---------------------------------------------------------------------------
const PAGE_HELPERS = `
window.__apaCheck = (function () {
  function visible(el) {
    if (!el) return false;
    if (typeof el.checkVisibility === 'function') {
      if (!el.checkVisibility({ checkOpacity: false, checkVisibilityCSS: true })) return false;
    }
    var r = el.getBoundingClientRect();
    return r.width > 0 || r.height > 0;
  }
  function rowShown(r) {
    if (r.hidden) return false;
    var cs = getComputedStyle(r);
    return cs.display !== 'none' && cs.visibility !== 'collapse' && cs.visibility !== 'hidden';
  }
  function table() { return document.querySelector('table[data-role="asset-table"]'); }
  function rows() {
    var t = table();
    if (!t) return [];
    return Array.prototype.filter.call(t.querySelectorAll('tr[data-role="asset-row"]'), rowShown);
  }
  function rowIds() { return rows().map(function (r) { return r.getAttribute('data-record-id'); }); }
  function cards() {
    var out = [];
    document.querySelectorAll('[data-role="kpi-card"]').forEach(function (card, ci) {
      var vals = card.querySelectorAll('[data-role="kpi-value"]');
      var cardText = (card.textContent || '').replace(/\\s+/g, ' ').trim();
      if (!vals.length) out.push({ card: ci, cardId: card.getAttribute('data-metric-id'), missing: true });
      vals.forEach(function (el) {
        var holder = el.closest('[data-metric-id]') || card;
        out.push({
          card: ci, cardId: card.getAttribute('data-metric-id'), metricId: holder.getAttribute('data-metric-id'),
          dataValue: el.getAttribute('data-value'), dataStatus: el.getAttribute('data-status'),
          text: (el.innerText || el.textContent || '').trim(), visible: visible(el), cardText: cardText
        });
      });
    });
    return out;
  }
  function getPath(obj, key) {
    var parts = String(key || '').split('.');
    var x = obj;
    for (var i = 0; i < parts.length; i++) {
      if (x === null || x === undefined) return undefined;
      x = x[parts[i]];
    }
    return x;
  }
  function charts(all) {
    return Array.prototype.map.call(document.querySelectorAll('[data-role="chart"]'), function (c) {
      var empty = c.querySelector('[data-role="chart-empty"]');
      var nums = Array.prototype.map.call(c.querySelectorAll('[data-role="chart-number"]'), function (n) {
        var key = n.getAttribute('data-key');
        var exp = all ? getPath(all, key) : undefined;
        return { key: key, dataValue: n.getAttribute('data-value'), text: (n.textContent || '').trim(),
          expected: typeof exp === 'number' ? exp : (exp === undefined ? 'missing' : null) };
      });
      return { id: c.getAttribute('data-chart-id'), svg: !!c.querySelector('svg'),
        empty: !!empty, emptyText: empty ? (empty.textContent || '').trim() : null, numbers: nums };
    });
  }
  /** The yield metric the asset table compares with (same choice as the template's comparisonMetric). */
  function comparisonMetric(all) {
    if (!all || !all.metrics) return null;
    var cfg = window.APA.data.config || {};
    var pref = (cfg.dashboard && cfg.dashboard.comparison_metric) ? [cfg.dashboard.comparison_metric] : [];
    pref = pref.concat(['gross_yield_investment', 'gross_yield_income_producing', 'gross_yield_owned']);
    for (var i = 0; i < pref.length; i++) if (all.metrics[pref[i]]) return all.metrics[pref[i]];
    var spec = (cfg.metrics || []).filter(function (m) { return m && m.kind === 'ratio' && m.num === '{REV}' && m.den === '{VAL}'; })[0];
    return spec ? all.metrics[spec.id] : null;
  }
  /** Value, revenue and yield cells of the visible rows, with the record values they must show. */
  function tableCells(filters, all) {
    var t = table();
    if (!t || !window.APA || typeof window.APA.resolveField !== 'function') return null;
    var idx = {};
    Array.prototype.forEach.call(t.querySelectorAll('th[data-sort]'), function (th) { idx[th.getAttribute('data-sort')] = th.cellIndex; });
    var vf = window.APA.resolveField('{VAL}', filters), rf = window.APA.resolveField('{REV}', filters);
    var KNOWN = ['value', 'derived', 'zero_recorded', 'zero_confirmed'];
    var recs = {};
    (window.APA.data.records || []).forEach(function (r) { recs[r.id] = r; });
    var fval = function (rec, f) {
      var tv = f && rec && rec.fields ? rec.fields[f] : null;
      return tv && KNOWN.indexOf(tv.s) >= 0 && typeof tv.v === 'number' ? tv.v : null;
    };
    var cell = function (row, key) { var i = idx[key]; return i !== undefined && row.cells[i] ? (row.cells[i].textContent || '').replace(/\\s+/g, ' ').trim() : null; };
    var cm = comparisonMetric(all);
    // outside the yield metric's valid periods the card says «غير قابل للمقارنة»: no row may show a yield number
    var yNc = !!(cm && cm.status === 'not_comparable');
    return { vf: vf, rf: rf, yNc: yNc, cols: Object.keys(idx), rows: rows().slice(0, ${MAX_TABLE_ROWS}).map(function (row) {
      var id = row.getAttribute('data-record-id');
      var rec = recs[id];
      return { id: id, v: fval(rec, vf), rv: fval(rec, rf), value: cell(row, 'value'), revenue: cell(row, 'revenue'), yld: cell(row, 'yield') };
    }) };
  }
  function scopeText() {
    var el = document.querySelector('[data-role="filter-scope"]');
    return el ? (el.textContent || '').replace(/\\s+/g, ' ').trim() : null;
  }
  function compact(all) {
    var m = {};
    Object.keys(all.metrics || {}).forEach(function (id) {
      var r = all.metrics[id];
      m[id] = { value: r.value === undefined ? null : r.value, status: r.status, n_included: r.n_included };
    });
    return m;
  }
  function cardMeta(all) {
    var m = {};
    Object.keys(all.metrics || {}).forEach(function (id) {
      var r = all.metrics[id];
      m[id] = { kind: r.kind, unit: r.unit, scope_text: r.scope_text || null, basis: r.basis || null, valuation_date: r.valuation_date || null,
        comp: r.comparability ? r.comparability.status || null : null, abs_change: typeof r.abs_change === 'number' ? r.abs_change : null };
    });
    return m;
  }
  function snapshot() {
    var out = { error: null };
    var all = null;
    try {
      out.filters = window.APA_UI.getFilters();
      all = window.APA.computeAll(out.filters);
      out.expected = compact(all);
      out.meta = cardMeta(all);
      out.expectedRows = window.APA.filterRecords(out.filters).length;
      out.table = tableCells(out.filters, all);
    } catch (e) { out.error = String(e && e.message || e); }
    out.rows = rowIds();
    out.cards = cards();
    out.charts = charts(all);
    out.scope = scopeText();
    out.selects = Array.prototype.map.call(document.querySelectorAll('select[data-filter]'), function (s) {
      return { dim: s.getAttribute('data-filter'), value: s.value };
    });
    return out;
  }
  function signature() {
    return JSON.stringify([rowIds(), scopeText(), cards().map(function (c) { return [c.dataValue, c.dataStatus, c.text]; }),
      document.querySelectorAll('[data-role="asset-detail"]').length]);
  }
  return { visible: visible, rows: rows, rowIds: rowIds, cards: cards, charts: charts, scopeText: scopeText,
    snapshot: snapshot, signature: signature, table: table, compact: compact, tableCells: tableCells };
})();
`;

async function installHelpers(page) {
  // normally present already through context.addInitScript; re-install defensively
  const ok = await page.evaluate(() => !!window.__apaCheck);
  if (!ok) await page.addScriptTag({ content: PAGE_HELPERS });
}

async function settle(page) {
  await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(r, 0)))));
}

/** Wait until the rendered state stops changing (handles debounced inputs), max ~2.5 s. */
async function waitStable(page) {
  await settle(page);
  let prev = await page.evaluate(() => window.__apaCheck.signature());
  let stableFor = 0;
  for (let i = 0; i < 25 && stableFor < 2; i++) {
    await page.waitForTimeout(100);
    const cur = await page.evaluate(() => window.__apaCheck.signature());
    if (cur === prev) stableFor++;
    else { stableFor = 0; prev = cur; }
  }
}

function compareCards(snap) {
  const problems = [];
  if (snap.error) return ['تعذر حساب القيم في الصفحة: ' + snap.error];
  for (const c of snap.cards) {
    if (c.missing) { problems.push(`بطاقة رقم ${c.card + 1} (${c.cardId}) بلا [data-role=kpi-value]`); continue; }
    const id = c.metricId;
    const exp = snap.expected[id];
    if (!id || !exp) { problems.push(`بطاقة ${id || '(بلا data-metric-id)'}: المؤشر غير موجود في computeAll`); continue; }
    if (c.dataValue === null) problems.push(`${id}: الخاصية data-value مفقودة`);
    else if (exp.value === null) {
      if (c.dataValue !== '') problems.push(`${id}: القيمة null لكن data-value="${c.dataValue}" (المتوقع "")`);
    } else {
      const got = c.dataValue === '' ? null : Number(c.dataValue);
      if (got === null || Number.isNaN(got) || !numClose(got, exp.value)) {
        problems.push(`${id}: data-value="${c.dataValue}" لا يطابق computeAll=${exp.value}`);
      }
    }
    if (c.dataStatus === null) problems.push(`${id}: الخاصية data-status مفقودة`);
    else if (c.dataStatus !== exp.status) problems.push(`${id}: data-status="${c.dataStatus}" والمتوقع "${exp.status}"`);
    if (!c.visible) problems.push(`${id}: قيمة البطاقة غير ظاهرة`);
    if (exp.value === null) {
      if (looksZero(c.text)) problems.push(`${id}: المؤشر غير قابل للاحتساب لكن البطاقة تعرض صفرًا «${c.text}»`);
      else if (!hasArabicWords(c.text)) problems.push(`${id}: المؤشر غير قابل للاحتساب ويجب عرض نص عربي يوضح ذلك، المعروض «${c.text}»`);
    } else if (!c.text) {
      problems.push(`${id}: البطاقة لا تعرض نصًا للقيمة`);
    } else if (snap.meta && snap.meta[id] && snap.meta[id].kind === 'growth' && snap.meta[id].comp !== 'confirmed') {
      // metrics.md §5: an unconfirmed growth shows its change in riyals, never a percentage
      const mt = snap.meta[id];
      if (/[\d٠-٩]\s*[%٪]/.test(c.text)) problems.push(`${id}: قابلية المقارنة غير مؤكدة لكن البطاقة تعرض نسبة «${short(c.text, 60)}» (المطلوب التغير بالريال)`);
      else if (mt.abs_change !== null && !shownMatches(c.text, mt.abs_change)) problems.push(`${id}: البطاقة «${short(c.text, 60)}» لا تعرض التغير بالريال ${mt.abs_change}`);
    } else if (!shownMatches(c.text, exp.value)) {
      // the number the board reads, not only the attribute (a formatting fault leaves data-value intact)
      problems.push(`${id}: الرقم المعروض «${short(c.text, 60)}» لا يطابق computeAll=${exp.value}`);
    }
    if (/NaN|undefined|Infinity|\[object/.test(c.text)) problems.push(`${id}: نص غير صالح في البطاقة «${c.text}»`);
    // brief §11: under a yield card, its scope, basis and valuation date — as computed for the current filters
    const meta = snap.meta && snap.meta[id];
    if (meta && meta.kind === 'ratio' && exp.value !== null && c.cardText !== undefined) {
      const norm = (s) => String(s).replace(/\s+/g, ' ').trim();
      for (const [k, label] of [['scope_text', 'النطاق'], ['basis', 'الأساس'], ['valuation_date', 'تاريخ التقييم']]) {
        if (meta[k] && !norm(c.cardText).includes(norm(meta[k]))) problems.push(`${id}: البطاقة لا تعرض ${label} المحسوب «${short(meta[k], 80)}»`);
      }
    }
  }
  return problems;
}

function compareCharts(snap) {
  const problems = [];
  for (const ch of snap.charts) {
    if (!ch.svg && !ch.empty) problems.push(`الرسم ${ch.id}: لا يحتوي <svg> ولا [data-role=chart-empty]`);
    else if (!ch.svg && ch.empty && !ch.emptyText) problems.push(`الرسم ${ch.id}: حالة فارغة دون سبب مكتوب`);
    // a drawn chart declares the numbers it prints, so they can be held against computeAll (template contract)
    if (ch.svg && !(ch.numbers || []).length) problems.push(`الرسم ${ch.id}: لا يعلن أرقامه المعروضة [data-role=chart-number] للمقارنة بـcomputeAll`);
    for (const n of ch.numbers || []) {
      if (n.expected === 'missing') { problems.push(`الرسم ${ch.id}: الرقم ${n.key} غير موجود في computeAll`); continue; }
      const dv = n.dataValue === '' || n.dataValue === null ? null : Number(n.dataValue);
      if (!numClose(dv, n.expected)) problems.push(`الرسم ${ch.id}: ${n.key} = ${n.dataValue} لا يطابق computeAll=${n.expected} للمرشحات الحالية`);
      else if (dv !== null && !shownMatches(n.text, dv)) problems.push(`الرسم ${ch.id}: النص المعروض «${short(n.text, 60)}» لا يطابق ${n.key}=${dv}`);
    }
  }
  return problems;
}

/** Table cells against the records: value, revenue and yield shown for each visible row of the current filters. */
function compareTable(snap) {
  const t = snap.table;
  if (snap.error || !t) return [];
  const problems = [];
  const checkMoney = (row, col, exp, label) => {
    const text = row[col];
    if (text === null || text === undefined) return;
    if (typeof exp === 'number') {
      if (!shownMatches(text, exp)) problems.push(`${row.id}: خانة «${label}» تعرض «${short(text, 50)}» والقيمة في السجل ${exp}`);
    } else if (!hasArabicWords(text) && !isMissingText(text)) {
      problems.push(`${row.id}: خانة «${label}» تعرض «${short(text, 50)}» والقيمة في السجل غير معروفة`);
    }
  };
  for (const row of t.rows) {
    if (t.vf) checkMoney(row, 'value', row.v, 'القيمة');
    if (t.rf) checkMoney(row, 'revenue', row.rv, 'الإيراد');
    const y = row.yld;
    if (t.yNc && y !== null && y !== undefined && leadingNumber(y)) {
      problems.push(`${row.id}: خانة العائد تعرض رقمًا «${y}» ومؤشر العائد غير قابل للمقارنة لهذه الفترة (التقييم لا يقابل إيرادها)`);
    } else if (y !== null && y !== undefined && leadingNumber(y)) {
      if (!(typeof row.v === 'number' && row.v > 0 && typeof row.rv === 'number')) problems.push(`${row.id}: خانة العائد تعرض رقمًا «${y}» دون إيراد وتقييم معروفين`);
      else if (!shownMatches(y, row.rv / row.v * 100)) problems.push(`${row.id}: خانة العائد «${y}» لا تساوي الإيراد ÷ التقييم × 100 (${(row.rv / row.v * 100).toFixed(4)})`);
    }
    if (problems.length >= 12) break;
  }
  return problems;
}

function compareRows(snap) {
  if (snap.error) return [];
  return snap.rows.length === snap.expectedRows ? [] :
    [`عدد صفوف الجدول الظاهرة ${snap.rows.length} لا يساوي APA.filterRecords = ${snap.expectedRows}`];
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
/** Node-side helpers on fixed cases (no browser): displayed-number parsing and the shared PII patterns. */
function selfTest() {
  const fails = [];
  const ok = (name, cond) => { if (!cond) fails.push(name); };
  ok('money', shownMatches('24,550,000 ر.س', 24550000.4));
  ok('money wrong x10', !shownMatches('245,500,000 ر.س', 24550000));
  ok('negative', shownMatches('-2,352,000 ر.س', -2352000));
  ok('unicode minus', shownMatches('−2,352,000', -2352000));
  ok('pct', shownMatches('6.65%', 6.6489));
  ok('pct off', !shownMatches('6.75%', 6.6489));
  ok('arabic digits', shownMatches('٢٤٬٥٥٠٬٠٠٠', 24550000));
  ok('isolates', shownMatches('⁦78.6%⁩', 78.57));
  ok('leading status', leadingNumber('غير متاح «6 محلات»') === null);
  ok('leading number', (leadingNumber('1,641,000 ر.س صفر مسجل') || {}).value === 1641000);
  ok('raw plain', rawIsPlainNumber('512,345,678', 512345678) && !rawIsPlainNumber('0501234567', 501234567));
  ok('nAr', nAr(1, 'a', 'b', 'c', 'd') === 'a' && nAr(5, 'a', 'b', 'c', 'd') === '5 c' && nAr(12, 'a', 'b', 'c', 'd') === '12 d');
  const pii = loadPiiPatterns();
  ok('pii patterns self-test: ' + pii.selfTest.join('؛ '), pii.selfTest.length === 0);
  ok('pii block', pii.hits('(050) 123-4567', 'block').includes('mobile') && pii.hits('011-4567890', 'block').includes('landline'));
  ok('pii review', pii.hits('1023456789', 'review').includes('national_id') && !pii.hits('1023456789', 'block').length);
  ok('pii arabic digits', pii.hits('٠٥٠١٢٣٤٥٦٧', 'block').includes('mobile'));
  for (const f of fails) console.error('فشل: ' + f);
  console.log(`check_dashboard --self-test: ${fails.length ? 'فشل ' + fails.length : 'ناجح'}`);
  return fails.length ? 1 : 0;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.selfTest) return selfTest();
  const htmlPath = path.resolve(args.html);
  const outDir = path.resolve(args.out);
  const screensDir = path.join(outDir, 'screens');
  fs.mkdirSync(screensDir, { recursive: true });
  const report = {
    contract_version: '1.0', html: htmlPath, checked_at: new Date().toISOString(), passed: false,
    summary: null, checks: [], warnings: [], screenshots: [], console_errors: [], external_requests: []
  };
  const add = (name, passed, detail, data) => {
    const c = { name, passed: !!passed, detail: detail || (passed ? 'ناجح' : 'فشل') };
    if (data !== undefined) c.data = data;
    report.checks.push(c);
    return !!passed;
  };
  /** Run one group of checks; an exception fails the group's checks that were not recorded yet. */
  const section = async (names, fn) => {
    const n0 = report.checks.length;
    try {
      await fn();
    } catch (e) {
      const done = new Set(report.checks.slice(n0).map((c) => c.name));
      for (const n of names) if (!done.has(n)) add(n, false, 'توقف هذا الفحص بسبب خطأ: ' + short(String((e && e.message) || e), 400));
    }
  };
  const finish = () => {
    const failed = report.checks.filter((c) => !c.passed);
    report.passed = report.checks.length > 0 && failed.length === 0;
    report.summary = { total: report.checks.length, passed: report.checks.length - failed.length, failed: failed.map((c) => c.name) };
    fs.writeFileSync(path.join(outDir, '08_dashboard_check.json'), JSON.stringify(report, null, 2) + '\n', 'utf8');
    console.log(`فحص الداشبورد في المتصفح: ${report.passed ? 'ناجح' : 'غير ناجح'} — ${report.summary.passed}/${report.summary.total} فحصًا`);
    for (const c of failed) console.log(`  [فشل] ${c.name}: ${short(c.detail, 400)}`);
    for (const w of report.warnings) console.log(`  [تنبيه] ${w.check}: ${short(w.detail, 400)}`);
    if (report.screenshots.length) console.log('  اللقطات: ' + report.screenshots.map((s) => path.join(outDir, s)).join('، '));
    console.log('  التقرير: ' + path.join(outDir, '08_dashboard_check.json'));
    return report.passed ? 0 : 1;
  };

  if (!fs.existsSync(htmlPath)) {
    add('load', false, 'ملف الداشبورد غير موجود: ' + htmlPath);
    return finish();
  }
  const pw = await loadPlaywright();
  if (!pw) {
    add('playwright', false, 'Playwright غير متاح في Node؛ لم يُختبر الداشبورد في متصفح. ثبّته (npm i -g playwright) ثم أعد الفحص، أو اذكر في التسليم أن الداشبورد بُني ولم يُختبر في متصفح.');
    return finish();
  }
  let browser;
  try {
    browser = await pw.chromium.launch();
  } catch (e) {
    add('playwright', false, 'تعذر تشغيل متصفح Chromium عبر Playwright؛ لم يُختبر الداشبورد في متصفح: ' + short(String(e.message || e), 300));
    return finish();
  }

  const url = pathToFileURL(htmlPath).href;
  const errors = report.console_errors;
  const external = report.external_requests;
  const watch = (page, label) => {
    page.on('console', (msg) => { if (msg.type() === 'error') errors.push({ page: label, type: 'console', text: short(msg.text(), 500) }); });
    page.on('pageerror', (err) => errors.push({ page: label, type: 'pageerror', text: short(String((err && err.message) || err), 500) }));
    page.on('request', (req) => { const u = req.url(); if (!/^(file|data|blob|about):/i.test(u)) external.push({ page: label, url: u }); });
    page.on('popup', (p) => watch(p, label + ':popup'));
    page.setDefaultTimeout(args.timeout);
    page.setDefaultNavigationTimeout(args.timeout);
  };
  const ready = async (page) => {
    await page.waitForFunction(() => !!(window.APA && window.APA_UI && typeof window.APA.computeAll === 'function'), null, { timeout: args.timeout });
    await installHelpers(page);
    await waitStable(page);
  };
  const open = async (context, label) => {
    const page = await context.newPage();
    watch(page, label);
    await page.goto(url, { waitUntil: 'load' });
    await ready(page);
    return page;
  };
  const snap = (page) => page.evaluate(() => window.__apaCheck.snapshot());

  try {
    const desktop = await browser.newContext({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1,
      locale: 'ar-SA', timezoneId: 'Asia/Riyadh', bypassCSP: true });
    await desktop.addInitScript({ content: PAGE_HELPERS });
    let page;
    try {
      page = await open(desktop, 'desktop');
      add('load', true, 'حُمّلت الصفحة وعُرّف window.APA وwindow.APA_UI');
    } catch (e) {
      add('load', false, 'تعذر تحميل الصفحة أو لم يُعرّف window.APA / window.APA_UI: ' + short(String(e.message || e), 400));
      add('console_errors', errors.length === 0, errors.length ? `${errors.length} خطأ: ` + short(errors.map((x) => x.text).join(' | '), 600) : 'لا أخطاء', errors.slice(0, 20));
      return finish();
    }

    // ---- draft builds are previews, never deliverables ----------------------------------------------
    await section(['draft'], async () => {
      const dr = await page.evaluate(() => {
        const m = (window.APA && window.APA.data && window.APA.data.meta) || {};
        const v = (window.APA && window.APA.data && window.APA.data.validation) || {};
        return { draft: m.draft === true, blocking: v.blocking_count, available: v.available };
      });
      if (!dr.draft) add('draft', true, 'نسخة نهائية (meta.draft = false)');
      else if (args.allowDraft) {
        add('draft', true, 'مسودة (meta.draft = true) فُحصت بالخيار --allow-draft؛ لا تُسلَّم قبل اجتياز التحقق وإعادة البناء دون --draft');
        report.warnings.push({ check: 'draft', detail: 'اللوحة مسودة لم تجتز التحقق الآلي؛ لا تُسلَّم' });
      } else {
        add('draft', false, 'اللوحة مسودة (meta.draft = true' + (dr.blocking ? `، أخطاء مانعة ${dr.blocking}` : (dr.available === false ? '، لا نتيجة تحقق' : '')) +
          ')؛ أصلح التحقق وأعد البناء دون --draft، أو مرّر --allow-draft لفحص مسودة لا تُسلَّم');
      }
    });

    // ---- desktop screenshots of the untouched default view: full page + viewport-sized crops ------------
    await section(['desktop_screenshot'], async () => {
      await page.screenshot({ path: path.join(screensDir, 'desktop.png'), fullPage: true });
      report.screenshots.push('screens/desktop.png');
      for (const f of fs.readdirSync(screensDir)) if (/^desktop_\d+\.png$/.test(f)) fs.unlinkSync(path.join(screensDir, f));
      const vp = page.viewportSize() || { width: 1440, height: 900 };
      const fullH = await page.evaluate(() => Math.max(document.documentElement.scrollHeight, document.body.scrollHeight));
      // a remainder shorter than MIN_TAIL px is folded into the last crop (never a 13 px sliver)
      const nFull = Math.max(1, Math.ceil((fullH - MIN_TAIL) / vp.height));
      const n = Math.min(MAX_CROPS, nFull);
      const crops = [];
      for (let i = 0; i < n; i++) {
        const name = `desktop_${String(i + 1).padStart(2, '0')}.png`;
        const y = i * vp.height;
        const hgt = (i === n - 1 && n === nFull) ? fullH - y : Math.min(vp.height, fullH - y);
        await page.screenshot({ path: path.join(screensDir, name), fullPage: true,
          clip: { x: 0, y, width: vp.width, height: hgt } });
        crops.push('screens/' + name);
      }
      report.screenshots.push(...crops);
      add('desktop_screenshot', true, `لقطة سطح المكتب 1440×900 (الصفحة كاملة): screens/desktop.png، و${nAr(crops.length, 'لقطة واحدة', 'لقطتان', 'لقطات', 'لقطة')} بحجم الشاشة ` +
        `(${crops[0] || '—'} …)` + (nFull > n ? ` — أول ${n} شاشات من ${nFull}` : ''));
    });

    // ---- html lang / dir, API, structure -----------------------------------------------------------
    let st = { reset: 0, table: 0, search: 0, sortable: 0, sections: [] };
    await section(['html_lang_dir', 'api', 'structure', 'details_sections'], async () => {
      const ld = await page.evaluate(() => ({
        lang: document.documentElement.getAttribute('lang'), dir: document.documentElement.getAttribute('dir'),
        body: getComputedStyle(document.body).direction,
        main: document.querySelector('main') ? getComputedStyle(document.querySelector('main')).direction : null
      }));
      const dirOk = ld.lang === 'ar' && ld.dir === 'rtl' && ld.body === 'rtl' && (ld.main === null || ld.main === 'rtl');
      add('html_lang_dir', dirOk, `lang="${ld.lang}" dir="${ld.dir}"، الاتجاه المحسوب: body=${ld.body}` + (ld.main ? `، main=${ld.main}` : '') +
        (dirOk ? '' : ' (المطلوب lang="ar" وdir="rtl" واتجاه محسوب rtl للصفحة والمحتوى)'));

      const api = await page.evaluate(() => {
        const A = window.APA, U = window.APA_UI;
        const fn = (o, k) => !!o && typeof o[k] === 'function';
        return { compute: fn(A, 'compute'), computeAll: fn(A, 'computeAll'), filterRecords: fn(A, 'filterRecords'),
          data: !!(A && A.data && A.data.metrics && A.data.config && Array.isArray(A.data.records)),
          getFilters: fn(U, 'getFilters'), setFilters: fn(U, 'setFilters'), render: fn(U, 'render') };
      });
      const apiMissing = Object.keys(api).filter((k) => !api[k]);
      add('api', apiMissing.length === 0, apiMissing.length ? 'عناصر ناقصة في واجهة APA/APA_UI: ' + apiMissing.join('، ')
        : 'APA.{compute, computeAll, filterRecords, data} وAPA_UI.{getFilters, setFilters, render} موجودة');

      st = await page.evaluate(() => {
        const q = (s) => document.querySelectorAll(s).length;
        return {
          cards: q('[data-role="kpi-card"]'), cardsNoId: q('[data-role="kpi-card"]:not([data-metric-id])'),
          table: q('table[data-role="asset-table"]'), search: q('input[data-role="table-search"]'),
          sortable: q('table[data-role="asset-table"] th[data-sort]'),
          reset: q('[data-action="reset-filters"]'), scope: q('[data-role="filter-scope"]'), print: q('[data-action="print"]'),
          filters: q('select[data-filter]'), period: q('select[data-filter="period"]'), charts: q('[data-role="chart"]'),
          topGap: q('[data-role="top-gap"]'), topAction: q('[data-role="top-action"]'),
          plan30: q('[data-role="plan-30"]'), plan60: q('[data-role="plan-60"]'), plan90: q('[data-role="plan-90"]'),
          sections: Array.prototype.map.call(document.querySelectorAll('details[data-role="details-section"]'), (d) => d.getAttribute('data-section'))
        };
      });
      const sp = [];
      if (st.cards < 1) sp.push('لا توجد بطاقات مؤشرات [data-role=kpi-card]');
      if (st.cards > MAX_CARDS) sp.push(`عدد البطاقات ${st.cards} يتجاوز ${MAX_CARDS}`);
      if (st.cardsNoId) sp.push(`${st.cardsNoId} بطاقة بلا data-metric-id`);
      for (const [k, label] of [['table', 'table[data-role=asset-table]'], ['search', 'input[data-role=table-search]'],
        ['sortable', 'th[data-sort]'], ['reset', '[data-action=reset-filters]'], ['scope', '[data-role=filter-scope]'],
        ['print', '[data-action=print]'], ['period', 'select[data-filter=period]'], ['charts', '[data-role=chart]'],
        ['plan30', '[data-role=plan-30]'], ['plan60', '[data-role=plan-60]'], ['plan90', '[data-role=plan-90]']]) {
        if (!st[k]) sp.push('عنصر مفقود: ' + label);
      }
      if (st.topGap > 3) sp.push(`عدد [data-role=top-gap] = ${st.topGap} (الحد 3)`);
      if (st.topAction > 3) sp.push(`عدد [data-role=top-action] = ${st.topAction} (الحد 3)`);
      add('structure', sp.length === 0, sp.length ? sp.join('؛ ')
        : `بطاقات ${st.cards}، رسوم ${st.charts}، مرشحات ${st.filters}، أهم الفجوات ${st.topGap}، أهم الإجراءات ${st.topAction}`, st);
      const missingSections = DETAIL_SECTIONS.filter((s) => !st.sections.includes(s));
      add('details_sections', missingSections.length === 0,
        missingSections.length ? 'أقسام تفاصيل مفقودة: ' + missingSections.join('، ') : 'جميع أقسام التفاصيل موجودة (' + DETAIL_SECTIONS.length + ')',
        { present: st.sections });
    });

    // ---- parity -----------------------------------------------------------------------------------
    await section(['parity'], async () => {
      const par = await page.evaluate(() => {
        const cases = (window.APA.data && window.APA.data.metrics && window.APA.data.metrics.parity_cases) || null;
        if (!Array.isArray(cases)) return { missing: true };
        return { cases: cases.map((c) => {
          try {
            return { filters: c.filters, expected: c.metrics, got: window.__apaCheck.compact(window.APA.computeAll(c.filters)) };
          } catch (e) { return { filters: c.filters, error: String((e && e.message) || e) }; }
        }) };
      });
      if (par.missing) { add('parity', false, 'لا توجد parity_cases في APA.data.metrics (05_metrics.json)'); return; }
      const mism = [];
      let compared = 0;
      for (const c of par.cases) {
        if (c.error) { mism.push({ filters: c.filters, error: c.error }); continue; }
        for (const id of Object.keys(c.expected || {})) {
          compared++;
          const e = c.expected[id], g = c.got[id];
          if (!g) { mism.push({ filters: c.filters, metric: id, problem: 'المؤشر غير موجود في computeAll' }); continue; }
          const bad = [];
          if (!numClose(e.value, g.value)) bad.push('value');
          if (e.status !== g.status) bad.push('status');
          if (e.n_included !== g.n_included) bad.push('n_included');
          if (bad.length) mism.push({ filters: c.filters, metric: id, fields: bad, python: e, js: g });
        }
      }
      add('parity', par.cases.length > 0 && mism.length === 0,
        par.cases.length === 0 ? 'parity_cases فارغة'
          : mism.length ? `${mism.length} عدم تطابق بين محرك الصفحة و05_metrics.json: ` + short(mism.slice(0, 3), 600)
            : `تطابق ${compared} قيمة في ${par.cases.length} حالة (value/status/n_included، تسامح نسبي ${PARITY_REL_TOL})`,
        mism.length ? mism.slice(0, 50) : undefined);
    });

    // ---- default state: cards, charts ------------------------------------------------------------------
    let def = null;
    await section(['cards', 'charts', 'table_cells'], async () => {
      def = await snap(page);
      const cardIds = await page.evaluate(() => ((window.APA.data.config && window.APA.data.config.metrics) || [])
        .filter((m) => m && m.card).map((m) => m.id));
      const cardProblems = compareCards(def);
      const shown = new Set(def.cards.map((c) => c.metricId));
      const notShown = cardIds.filter((id) => !shown.has(id));
      if (notShown.length) cardProblems.push('مؤشرات card:true غير معروضة في البطاقات: ' + notShown.join('، '));
      add('cards', cardProblems.length === 0 && def.cards.length > 0,
        cardProblems.length ? cardProblems.join('؛ ') : `${nAr(def.cards.length, 'قيمة بطاقة واحدة', 'قيمتا بطاقة', 'قيم بطاقات', 'قيمة بطاقة')} تطابق computeAll للمرشحات الافتراضية (الخاصية والرقم المعروض)`,
        { filters: def.filters, cards: def.cards.map((c) => ({ id: c.metricId, value: c.dataValue, status: c.dataStatus, text: c.text })) });
      const chartProblems = compareCharts(def);
      const nNums = def.charts.reduce((n, c) => n + (c.numbers || []).length, 0);
      add('charts', chartProblems.length === 0 && def.charts.length > 0,
        def.charts.length === 0 ? 'لا توجد رسوم [data-role=chart]'
          : chartProblems.length ? chartProblems.join('؛ ')
            : `${nAr(def.charts.length, 'رسم واحد', 'رسمان', 'رسوم', 'رسمًا')}: كل رسم يحتوي SVG أو حالة فارغة مسببة، و${nAr(nNums, 'رقم واحد', 'رقمان', 'أرقام', 'رقمًا')} معروضة في الرسوم تطابق computeAll`);
      const tableProblems = compareTable(def);
      add('table_cells', !!def.table && tableProblems.length === 0,
        !def.table ? 'تعذر قراءة خانات الجدول (الجدول أو APA.resolveField غير موجود)'
          : tableProblems.length ? tableProblems.join('؛ ')
            : `خانات القيمة والإيراد والعائد في ${def.table.rows.length} صفًا تطابق السجلات (الحقلان ${def.table.vf || '—'} و${def.table.rf || '—'})`);
    });

    // ---- filters ------------------------------------------------------------------------------------------
    const resetFilters = async (probs, label) => {
      const btn = page.locator('[data-action="reset-filters"]').first();
      if (st.reset && (await btn.isVisible().catch(() => false))) {
        await btn.click();
      } else {
        if (probs && label) probs.push(`${label}: زر إعادة الضبط [data-action=reset-filters] غير ظاهر`);
        await page.evaluate((f) => { window.APA_UI.setFilters(f); window.APA_UI.render(); }, def ? def.filters : {});
      }
      await waitStable(page);
    };
    const checkDefaultRestored = async (label, probs) => {
      const s = await snap(page);
      if (!sameJSON(cleanFilters(s.filters), cleanFilters(def.filters))) probs.push(`${label}: لم تُستعد المرشحات الافتراضية (${short(s.filters)})`);
      if (s.scope !== def.scope) probs.push(`${label}: نص نطاق التصفية لم يعد إلى الافتراضي`);
      if (s.rows.length !== def.rows.length) probs.push(`${label}: عدد الصفوف ${s.rows.length} بدل ${def.rows.length}`);
      if (def.cards.map((c) => c.dataValue).join('|') !== s.cards.map((c) => c.dataValue).join('|')) probs.push(`${label}: قيم البطاقات لم تعد إلى الافتراضي`);
    };
    let selects = [];
    await section(['filters'], async () => {
      if (!def) def = await snap(page);
      selects = await page.evaluate(() => Array.prototype.map.call(document.querySelectorAll('select[data-filter]'), (s) => ({
        dim: s.getAttribute('data-filter'),
        options: Array.prototype.filter.call(s.options, (o) => !o.disabled).map((o) => o.value)
      })));
      const configured = await page.evaluate(() => (window.APA.data.config && window.APA.data.config.filters) || []);
      const fProblems = [];
      const tried = [];
      for (const dim of configured) if (!selects.some((s) => s.dim === dim)) fProblems.push(`المرشح المعرّف في الإعدادات «${dim}» لا يقابله select[data-filter="${dim}"]`);
      for (const sel of selects.filter((s) => s.dim !== 'period')) {
        if (!sel.options.includes('')) fProblems.push(`المرشح ${sel.dim}: لا يوجد خيار «الكل» بقيمة ""`);
        const values = sel.options.filter((v) => v !== '').slice(0, MAX_FILTER_VALUES);
        if (!values.length) fProblems.push(`المرشح ${sel.dim}: لا توجد خيارات`);
        for (const v of values) {
          await resetFilters();
          await page.locator(`select[data-filter="${sel.dim}"]`).first().selectOption(v);
          await waitStable(page);
          const s = await snap(page);
          const label = `${sel.dim}=${v}`;
          tried.push({ filter: label, rows: s.rows.length, expectedRows: s.expectedRows });
          if (s.error) { fProblems.push(`${label}: ${s.error}`); continue; }
          if (String(s.filters[sel.dim]) !== v) fProblems.push(`${label}: APA_UI.getFilters().${sel.dim} = ${short(s.filters[sel.dim])}`);
          compareRows(s).forEach((p) => fProblems.push(`${label}: ${p}`));
          compareCards(s).forEach((p) => fProblems.push(`${label}: ${p}`));
          compareCharts(s).forEach((p) => fProblems.push(`${label}: ${p}`));
          compareTable(s).forEach((p) => fProblems.push(`${label}: ${p}`));
          if (s.scope === def.scope) fProblems.push(`${label}: نص [data-role=filter-scope] لم يتغير («${short(s.scope, 80)}»)`);
          await resetFilters(fProblems, `إعادة الضبط بعد ${label}`);
          await checkDefaultRestored(`إعادة الضبط بعد ${label}`, fProblems);
        }
      }
      await page.evaluate((f) => { window.APA_UI.setFilters(f); window.APA_UI.render(); }, def.filters);
      await waitStable(page);
      await checkDefaultRestored('APA_UI.setFilters(الافتراضي)', fProblems);
      const dimSelects = selects.filter((s) => s.dim !== 'period').length;
      add('filters', fProblems.length === 0,
        fProblems.length ? fProblems.slice(0, 12).join('؛ ')
          : dimSelects === 0 ? 'لا توجد مرشحات أبعاد معرّفة في الإعدادات (يوجد مرشح الفترة فقط)'
            : `جُرّب ${nAr(tried.length, 'خيار واحد', 'خياران', 'خيارات', 'خيارًا')} في ${nAr(dimSelects, 'مرشح واحد', 'مرشحين', 'مرشحات', 'مرشحًا')}: الصفوف والبطاقات والرسوم وخانات الجدول ونص النطاق متسقة، وإعادة الضبط تستعيد الافتراضي`,
        { tried, problems: fProblems.slice(0, 50) });
    });

    // ---- period filter -------------------------------------------------------------------------------------
    await section(['period_filter'], async () => {
      const pProblems = [];
      const periodSel = selects.find((s) => s.dim === 'period');
      const periodsTried = [];
      if (!periodSel) pProblems.push('مرشح الفترة select[data-filter=period] غير موجود');
      else {
        for (const v of periodSel.options.filter((x) => x !== '')) {
          await page.locator('select[data-filter="period"]').first().selectOption(v);
          await waitStable(page);
          const s = await snap(page);
          periodsTried.push(v);
          if (s.error) { pProblems.push(`period=${v}: ${s.error}`); continue; }
          if (String(s.filters.period) !== v) pProblems.push(`period=${v}: APA_UI.getFilters().period = ${short(s.filters.period)}`);
          compareRows(s).forEach((p) => pProblems.push(`period=${v}: ${p}`));
          compareCards(s).forEach((p) => pProblems.push(`period=${v}: ${p}`));
          compareCharts(s).forEach((p) => pProblems.push(`period=${v}: ${p}`));
          compareTable(s).forEach((p) => pProblems.push(`period=${v}: ${p}`));
        }
        await resetFilters(periodsTried.length > 1 ? pProblems : null, 'إعادة الضبط بعد تغيير الفترة');
        await checkDefaultRestored('إعادة الضبط بعد تغيير الفترة', pProblems);
      }
      add('period_filter', pProblems.length === 0 && periodsTried.length > 0,
        pProblems.length ? pProblems.slice(0, 12).join('؛ ') : `جُرّبت الفترات: ${periodsTried.join('، ')}`);
    });

    // ---- table: search -----------------------------------------------------------------------------------
    await section(['table_search'], async () => {
      await page.evaluate(() => {
        const t = window.__apaCheck.table();
        for (let el = t && t.parentElement; el; el = el.parentElement) if (el.tagName === 'DETAILS') el.open = true;
      });
      await waitStable(page);
      const tProblems = [];
      const before = await page.evaluate(() => window.__apaCheck.rowIds());
      const firstId = before[0];
      let term = null;
      if (!st.table || !st.search) tProblems.push('الجدول أو حقل البحث غير موجود');
      else if (!before.length) tProblems.push('لا توجد صفوف ظاهرة في الجدول');
      else {
        const name = await page.evaluate((id) => {
          const r = (window.APA.data.records || []).find((x) => x.id === id);
          if (r && typeof r.name === 'string' && r.name.trim()) return r.name.trim();
          const row = document.querySelector('table[data-role="asset-table"] tr[data-role="asset-row"][data-record-id="' + CSS.escape(id) + '"]');
          const cell = row && row.querySelector('td');
          return cell ? (cell.textContent || '').trim() : '';
        }, firstId);
        const words = name.split(/\s+/).filter((w) => w.length >= 2).sort((a, b) => b.length - a.length);
        term = name.length <= 4 || !words.length ? name : words[0];
        const search = page.locator('input[data-role="table-search"]').first();
        const type = async (text) => { await search.fill(text); await search.dispatchEvent('change'); await waitStable(page); };
        if (term) {
          await type(term);
          const after = await page.evaluate(() => window.__apaCheck.rowIds());
          if (!after.length) tProblems.push(`البحث عن «${term}» (من اسم أول صف) أخفى كل الصفوف`);
          if (after.length > before.length) tProblems.push(`البحث زاد عدد الصفوف (${before.length} → ${after.length})`);
          if (!after.includes(firstId)) tProblems.push(`البحث عن «${term}» أخفى الصف ${firstId} الذي يحتويه`);
        }
        await type('zqxj§لا_يوجد_هذا_النص§');
        const none = await page.evaluate(() => window.__apaCheck.rowIds());
        if (none.length) tProblems.push(`البحث عن نص غير موجود أبقى ${none.length} صفًا`);
        await type('');
        const restored = await page.evaluate(() => window.__apaCheck.rowIds());
        if (!sameJSON([...restored].sort(), [...before].sort())) tProblems.push(`مسح البحث لم يستعد الصفوف (${restored.length} بدل ${before.length})`);
      }
      add('table_search', tProblems.length === 0, tProblems.length ? tProblems.join('؛ ')
        : `البحث عن «${term}» أبقى الصف المطلوب ولم يزد الصفوف، ونص غير موجود أخفاها كلها، والمسح أعادها (${before.length})`);
    });

    // ---- table: sort -----------------------------------------------------------------------------------
    await section(['table_sort'], async () => {
      const sProblems = [];
      const sortInfo = [];
      const nHeaders = Math.min(st.sortable, MAX_SORT_HEADERS);
      const header = (i) => page.locator('table[data-role="asset-table"] th[data-sort]').nth(i);
      for (let i = 0; i < nHeaders; i++) {
        const key = await header(i).getAttribute('data-sort');
        if (!(await header(i).isVisible())) { sortInfo.push({ key, skipped: 'hidden' }); continue; }
        const idsBefore = await page.evaluate(() => window.__apaCheck.rowIds());
        const readCol = (idx) => page.evaluate((k) => {
          const ths = document.querySelectorAll('table[data-role="asset-table"] th[data-sort]');
          const th = ths[k];
          const ci = th ? th.cellIndex : -1;
          const rows = window.__apaCheck.rows();
          return { numeric: !!th && th.classList.contains('n'), ids: rows.map((r) => r.getAttribute('data-record-id')),
            texts: rows.map((r) => (r.cells[ci] ? (r.cells[ci].textContent || '').trim() : null)) };
        }, idx);
        await header(i).click();
        await waitStable(page);
        const s1 = await readCol(i);
        await header(i).click();
        await waitStable(page);
        const s2 = await readCol(i);
        const ids2 = s2.ids;
        // Rows without a value ("—", "غير متاح", ...) may legitimately stay last in both directions:
        // only the rows holding a value must reverse, and only when they hold >= 2 distinct values.
        const textOf = new Map(s1.ids.map((id, k) => [id, s1.texts[k]]));
        const hasValue = (id) => { const t = textOf.get(id); return t !== null && t !== undefined && !isMissingText(t); };
        const valued1 = s1.ids.filter(hasValue), valued2 = ids2.filter(hasValue);
        const distinct = new Set(valued1.map((id) => textOf.get(id))).size;
        const changed = !sameJSON(valued1, valued2);
        // numeric columns: the shown numbers must be ordered, one direction per click
        let order = null;
        if (s1.numeric) {
          const nums = (st) => st.texts.map((t) => (t && !isMissingText(t) ? leadingNumber(t) : null)).filter(Boolean).map((x) => x.value);
          const dir = (a) => {
            let up = true, down = true;
            for (let k = 1; k < a.length; k++) { if (a[k] < a[k - 1]) up = false; if (a[k] > a[k - 1]) down = false; }
            return up && down ? 'flat' : up ? 'asc' : down ? 'desc' : 'none';
          };
          order = [dir(nums(s1)), dir(nums(s2))];
        }
        sortInfo.push({ key, rows: ids2.length, rows_with_value: valued1.length, distinct_values: distinct, changed, order });
        if (!sameJSON([...s1.ids].sort(), [...idsBefore].sort()) || !sameJSON([...ids2].sort(), [...idsBefore].sort())) {
          sProblems.push(`الفرز حسب ${key} غيّر مجموعة الصفوف`);
        } else if (distinct >= 2 && !changed) {
          sProblems.push(`النقر مرتين على عمود ${key} لم يعكس ترتيب الصفوف رغم اختلاف القيم`);
        } else if (order && (order.includes('none') || (order[0] !== 'flat' && order[0] === order[1]))) {
          sProblems.push(`الفرز حسب العمود الرقمي ${key} لا يرتب الأرقام المعروضة (${order.join(' ثم ')})`);
        }
      }
      if (!nHeaders) sProblems.push('لا توجد أعمدة قابلة للفرز th[data-sort]');
      add('table_sort', sProblems.length === 0, sProblems.length ? sProblems.join('؛ ') : `جُرّب الفرز (نقرتان) على ${nAr(sortInfo.length, 'عمود واحد', 'عمودين', 'أعمدة', 'عمودًا')}، والأعمدة الرقمية مرتبة`, sortInfo);
    });

    // ---- table: detail ---------------------------------------------------------------------------------
    await section(['table_detail'], async () => {
      const dProblems = [];
      const detailVisible = () => page.evaluate(() => Array.prototype.some.call(document.querySelectorAll('[data-role="asset-detail"]'), (d) => window.__apaCheck.visible(d)));
      const closeDetail = async () => {
        const btn = page.locator('[data-action="close-detail"]');
        const n = await btn.count();
        let done = false;
        for (let i = 0; i < n && !done; i++) {
          if (await btn.nth(i).isVisible()) { await btn.nth(i).click(); done = true; }
        }
        if (!done) dProblems.push('لا يوجد زر إغلاق ظاهر [data-action=close-detail] في لوحة التفاصيل');
        await waitStable(page);
      };
      const firstRow = page.locator('table[data-role="asset-table"] tr[data-role="asset-row"]:visible').first();
      if (!(await firstRow.count())) dProblems.push('لا يوجد صف ظاهر لفتح التفاصيل');
      else {
        if (await detailVisible()) dProblems.push('لوحة التفاصيل ظاهرة قبل اختيار أي صف');
        const rid = await firstRow.getAttribute('data-record-id');
        await firstRow.click();
        await waitStable(page);
        if (!(await detailVisible())) dProblems.push('النقر على الصف لم يُظهر [data-role=asset-detail]');
        else {
          const name = await page.evaluate((id) => { const r = (window.APA.data.records || []).find((x) => x.id === id); return r ? r.name : null; }, rid);
          const txt = await page.evaluate(() => Array.prototype.map.call(document.querySelectorAll('[data-role="asset-detail"]'), (d) => d.textContent || '').join(' '));
          if (name && !txt.includes(String(name).trim())) dProblems.push(`لوحة التفاصيل لا تعرض اسم السجل ${rid}`);
          await closeDetail();
          if (await detailVisible()) dProblems.push('زر الإغلاق [data-action=close-detail] لم يُخفِ لوحة التفاصيل');
        }
        await page.locator('table[data-role="asset-table"] tr[data-role="asset-row"]:visible').first().focus();
        await page.keyboard.press('Enter');
        await waitStable(page);
        if (!(await detailVisible())) dProblems.push('الضغط على Enter في الصف لم يُظهر التفاصيل (يلزم tabindex ومعالج لوحة المفاتيح)');
        else {
          await closeDetail();
          if (await detailVisible()) dProblems.push('لم تُغلق لوحة التفاصيل بعد فتحها بلوحة المفاتيح');
        }
      }
      add('table_detail', dProblems.length === 0, dProblems.length ? dProblems.join('؛ ') : 'النقر وEnter يفتحان التفاصيل ويُعرض اسم السجل، وزر الإغلاق يخفيها');
    });

    // ---- text checks with every <details> opened ----------------------------------------------------
    await section(['no_raw_tokens', 'privacy'], async () => {
      await page.evaluate(() => { document.querySelectorAll('details').forEach((d) => { d.open = true; }); });
      await waitStable(page);
      // a capture of the study details as the board would open them (C21: the full page alone is unreadable)
      try {
        const det = page.locator('#sec-details').first();
        if (await det.count()) {
          // the sticky filter bar would be painted over the capture of a section taller than the window
          const unstick = await page.addStyleTag({ content: '.filterbar{position:static!important}' });
          await det.screenshot({ path: path.join(screensDir, 'details_open.png') });
          report.screenshots.push('screens/details_open.png');
          // and readable slices of at most 1400 px (a 40,000 px image is shrunk past legibility when viewed)
          const box = await det.boundingBox();
          if (box && box.height > 0) {
            const top = box.y + (await page.evaluate(() => window.scrollY));
            const nAll = Math.max(1, Math.ceil((box.height - MIN_TAIL) / 1400));
            const n = Math.min(30, nAll);
            for (let i = 0; i < n; i++) {
              const y = top + i * 1400;
              const hgt = (i === n - 1 && n === nAll) ? box.height - i * 1400 : Math.min(1400, box.height - i * 1400);
              if (hgt <= 0) break;
              const name = `details_open_${String(i + 1).padStart(2, '0')}.png`;
              await page.screenshot({ path: path.join(screensDir, name), fullPage: true,
                clip: { x: Math.max(0, box.x), y, width: box.width, height: hgt } });
              report.screenshots.push('screens/' + name);
            }
          }
          await unstick.evaluate((el) => el.remove());
        }
      } catch (e) { report.warnings.push({ check: 'details_open', detail: 'تعذر التقاط screens/details_open.png: ' + short(String(e.message || e), 160) }); }
      const texts = await page.evaluate(() => {
        const out = [];
        const skip = { SCRIPT: 1, STYLE: 1, NOSCRIPT: 1, TEMPLATE: 1 };
        const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
          acceptNode: (n) => { for (let p = n.parentNode; p; p = p.parentNode) if (skip[p.nodeName]) return NodeFilter.FILTER_REJECT; return NodeFilter.FILTER_ACCEPT; }
        });
        for (let n = w.nextNode(); n; n = w.nextNode()) { const t = n.nodeValue; if (t && t.trim()) out.push(t); }
        out.push(document.title || '');
        const attrs = ['title', 'aria-label', 'alt', 'placeholder', 'aria-description', 'data-tooltip', 'data-label'];
        document.querySelectorAll('*').forEach((el) => {
          attrs.forEach((a) => { const v = el.getAttribute(a); if (v) out.push(v); });
          if ((el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') && el.value) out.push(el.value);
        });
        return out;
      });
      const dataStrings = await page.evaluate(() => {
        const out = [];
        const walk = (x, key, parent) => {
          if (typeof x === 'string') {
            if (key === 'sha256' || key === 'sha256_short' || /^[0-9a-f]{16,}$/i.test(x)) return;
            out.push({ s: x, rawOf: key === 'raw' && parent && typeof parent.v === 'number' ? parent.v : null });
          } else if (Array.isArray(x)) x.forEach((y) => walk(y, key, parent));
          else if (x && typeof x === 'object') Object.keys(x).forEach((k) => walk(x[k], k, x));
        };
        walk(window.APA.data, null, null);
        return out;
      });
      const BAD = /\bNaN\b|\bundefined\b|\[object Object\]|\bInfinity\b|\{\{\s*[mf]:/;
      const badHits = texts.filter((t) => BAD.test(t)).map((t) => short(t.trim(), 120));
      add('no_raw_tokens', badHits.length === 0, badHits.length
        ? 'نصوص غير صالحة في الصفحة (NaN/undefined/[object Object]/Infinity أو رموز {{m:}}/{{f:}} غير مُعالجة): ' + short(badHits.slice(0, 5), 500)
        : 'لا توجد قيم خام غير صالحة في نص الصفحة', badHits.length ? badHits.slice(0, 20) : undefined);
      // the shared patterns of assets/pii_patterns.json (same as apa_common.mask_pii), tested on ASCII digits
      let pii;
      try { pii = loadPiiPatterns(); } catch (e) {
        add('privacy', false, 'تعذر تحميل أنماط البيانات الشخصية ' + PII_PATTERNS_PATH + ': ' + short(String(e.message || e), 200));
        return;
      }
      const blockHits = [], reviewHits = [];
      const scan = (arr, where) => {
        for (const x of arr) {
          const t0 = typeof x === 'string' ? x : x.s;
          if (typeof x === 'object' && x.rawOf !== null && rawIsPlainNumber(t0, x.rawOf)) continue;
          for (const kind of pii.hits(t0, 'block')) blockHits.push({ where, kind });
          for (const kind of pii.hits(t0, 'review')) reviewHits.push({ where, kind });
        }
      };
      scan(texts, 'page_text');
      scan(dataStrings, 'embedded_data');
      const tally = (hits) => hits.reduce((a, x) => { const k = x.where + ':' + x.kind; a[k] = (a[k] || 0) + 1; return a; }, {});
      const fmtTally = (t) => Object.entries(t).map(([k, n]) => `${k}×${n}`).join('، ');
      const blockSum = tally(blockHits), reviewSum = tally(reviewHits);
      const probs = [];
      if (pii.selfTest.length) probs.push('أنماط pii_patterns.json لا تعمل في JS كما في Python: ' + pii.selfTest.slice(0, 5).join('؛ '));
      if (blockHits.length) probs.push('وُجدت بيانات شخصية (لا تُطبع القيم): ' + fmtTally(blockSum));
      if (reviewHits.length) {
        report.warnings.push({ check: 'privacy', detail: 'أرقام من 10 خانات تبدأ بـ1 أو 2 قد تكون أرقام هوية (أو صكوكًا) — راجعها: ' + fmtTally(reviewSum) });
      }
      add('privacy', probs.length === 0,
        probs.length ? probs.join('؛ ')
          : `لا جوال ولا هاتف ثابت ولا بريد ولا آيبان (أنماط pii_patterns.json الإصدار ${pii.version}) في نص الصفحة (${texts.length} نصًا) ولا في البيانات المضمّنة (${dataStrings.length} نصًا)` +
            (reviewHits.length ? `؛ للمراجعة: ${reviewHits.length} رقمًا قد يكون رقم هوية (انظر warnings)` : ''),
        blockHits.length || reviewHits.length ? { block: blockSum, review: reviewSum } : undefined);
    });

    // ---- with the details still open: growth shown as a %, raw engine keys, «Sheet!Cell» order ----------------
    await section(['growth_no_pct', 'no_raw_keys', 'bidi_refs'], async () => {
      // the asset detail panel too (its field notes and references are board-facing)
      await page.evaluate(() => { const row = document.querySelector('tr[data-role="asset-row"]'); if (row) row.click(); });
      await waitStable(page);
      const r = await page.evaluate(() => {
        const out = { growth: [], keys: [], bidi: { checked: 0, bad: [] } };
        const visibleEl = (el) => !!el && el.getClientRects().length > 0 && getComputedStyle(el).visibility !== 'hidden';
        // (1) metrics.md §5: a growth result without confirmed comparability never shows a percentage
        let all = null;
        try { all = window.APA.computeAll(window.APA_UI.getFilters()); } catch (e) { all = null; }
        Object.keys((all && all.metrics) || {}).forEach((id) => {
          const m = all.metrics[id];
          if (!m || m.kind !== 'growth' || (m.comparability && m.comparability.status === 'confirmed')) return;
          const sel = `[data-metric-id="${id}"] [data-role="kpi-value"], [data-metric-id="${id}"] [data-role="calc-value"], [data-metric-ref="${id}"]`;
          document.querySelectorAll(sel).forEach((el) => {
            const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
            if (/[\d٠-٩]\s*[%٪]/.test(t)) out.growth.push(id + ': «' + t.slice(0, 90) + '»');
          });
        });
        // (2) engine keys (rev.2025.excl_vat, zero_as_missing) in visible text outside <code> / .tech
        const skip = { SCRIPT: 1, STYLE: 1, NOSCRIPT: 1, TEMPLATE: 1, CODE: 1 };
        const KEY = /[a-z][a-z0-9]*(?:[._][a-z0-9]+)+/g;
        const FILE = /\.(xlsx|xlsm|xls|ods|csv|json|html?|md|pdf|png)$/i;
        const WEB = /^www\.|\.(com|sa|org|net|gov|edu)(\.|$)/i;
        const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
          acceptNode: (n) => {
            for (let p = n.parentNode; p && p !== document.body; p = p.parentNode) {
              if (skip[p.nodeName] || (p.classList && p.classList.contains('tech'))) return NodeFilter.FILTER_REJECT;
            }
            return NodeFilter.FILTER_ACCEPT;
          }
        });
        const nodes = [];
        for (let n = w.nextNode(); n; n = w.nextNode()) if (n.nodeValue && n.nodeValue.trim() && visibleEl(n.parentElement)) nodes.push(n);
        nodes.forEach((n) => {
          (n.nodeValue.match(KEY) || []).forEach((k) => { if (!FILE.test(k) && !WEB.test(k) && out.keys.length < 40) out.keys.push(k); });
        });
        // (3) «<Arabic>!<Cell>»: the «!» glyph must sit between the sheet name's last letter and the cell (dashboard-spec §7)
        const blockOf = (el) => { for (let p = el; p; p = p.parentElement) { const d = getComputedStyle(p).display; if (d !== 'inline' && d !== 'contents') return p; } return document.body; };
        const groups = new Map();
        nodes.forEach((n) => { const b = blockOf(n.parentElement); if (!groups.has(b)) groups.set(b, []); groups.get(b).push(n); });
        const rectOf = (node, i) => { const rg = document.createRange(); rg.setStart(node, i); rg.setEnd(node, i + 1); const rr = rg.getBoundingClientRect(); return rr; };
        const AR = /[\u0600-\u06FF]/;
        groups.forEach((ns) => {
          const chars = [];
          ns.forEach((n) => { for (let i = 0; i < n.nodeValue.length; i++) chars.push([n, i, n.nodeValue[i]]); });
          for (let i = 1; i < chars.length - 2 && out.bidi.checked < 3000; i++) {
            if (chars[i][2] !== '!' || !/[$A-Z]/.test(chars[i + 1][2])) continue;
            let j = i - 1;
            while (j >= 0 && /\s/.test(chars[j][2])) j--;
            if (j < 0 || !AR.test(chars[j][2])) continue;
            const a = rectOf(chars[j][0], chars[j][1]), b = rectOf(chars[i][0], chars[i][1]), c = rectOf(chars[i + 1][0], chars[i + 1][1]);
            if (!a.width || !b.width || !c.width) continue;
            const line = (x) => x.top + x.height / 2;
            if (Math.abs(line(a) - line(b)) > b.height / 2 || Math.abs(line(c) - line(b)) > b.height / 2) continue;   // wrapped
            out.bidi.checked++;
            const ax = a.left + a.width / 2, bx = b.left + b.width / 2, cx = c.left + c.width / 2;
            if (!(bx > Math.min(ax, cx) && bx < Math.max(ax, cx))) {
              let k = i + 1, cell = '';
              while (k < chars.length && /[$A-Z0-9:]/.test(chars[k][2]) && cell.length < 12) { cell += chars[k][2]; k++; }
              if (out.bidi.bad.length < 20) out.bidi.bad.push('…' + chars.slice(Math.max(0, j - 6), j + 1).map((x) => x[2]).join('') + '!' + cell);
            }
          }
        });
        return out;
      });
      add('growth_no_pct', r.growth.length === 0, r.growth.length
        ? 'مؤشر تغير بقابلية مقارنة غير مؤكدة يُعرض نسبةً (المطلوب التغير بالريال دون نسبة، metrics.md §5): ' + short(r.growth.slice(0, 4), 400)
        : 'لا نسبة نمو معروضة لمؤشر تغير غير مؤكد المقارنة (البطاقات وسجل الحسابات ورقاقات الدليل)');
      add('no_raw_keys', r.keys.length === 0, r.keys.length
        ? 'مفاتيح تقنية إنجليزية في النص الظاهر (تُعرض بالتسمية العربية، والمفتاح في <code> أو «المرجع التقني» فقط): ' + short([...new Set(r.keys)].slice(0, 12), 400)
        : 'لا مفاتيح حقول أو مؤشرات أو حالات إنجليزية في النص الظاهر خارج المراجع التقنية');
      add('bidi_refs', r.bidi.bad.length === 0, r.bidi.bad.length
        ? `${r.bidi.bad.length} مرجع خلية بعد اسم ورقة عربي تظهر فيه «!» خارج موضعها (الخلية ملتصقة بالاسم): ` + short(r.bidi.bad.slice(0, 6), 400)
        : `«!» بين اسم الورقة والخلية في ${r.bidi.checked} مرجعًا ظاهرًا بعد اسم ورقة عربي`);
    });

    // ---- every visible button / [data-action] -----------------------------------------------------------
    await section(['buttons'], async () => {
      await page.reload({ waitUntil: 'load' });
      await ready(page);
      const stubPrint = () => page.evaluate(() => { window.__apaPrintCalls = 0; window.print = function () { window.__apaPrintCalls++; }; });
      await stubPrint();
      let navigated = false;
      page.on('framenavigated', (fr) => { if (fr === page.mainFrame()) navigated = true; });
      const bProblems = [];
      const clicked = [];
      const sel = 'button, [data-action]';
      const nButtons = Math.min(await page.locator(sel).count(), MAX_BUTTONS);
      for (let i = 0; i < nButtons; i++) {
        const loc = page.locator(sel).nth(i);
        if (!(await loc.count()) || !(await loc.isVisible().catch(() => false))) continue;
        const info = await loc.evaluate((el) => ({ action: el.getAttribute('data-action'), text: (el.textContent || '').trim().slice(0, 40),
          tag: el.tagName, href: el.getAttribute('href') }));
        if (info.tag === 'A' && info.href && !info.href.startsWith('#')) continue;
        const what = info.action || info.text || info.tag;
        const errBefore = errors.length;
        const printsBefore = await page.evaluate(() => window.__apaPrintCalls);
        navigated = false;
        try {
          await loc.click({ timeout: 3000 });
        } catch (e) {
          try { await loc.evaluate((el) => el.click()); } catch (e2) { bProblems.push(`تعذر النقر على «${what}»: ${short(String(e2.message || e2), 120)}`); }
        }
        await settle(page);
        clicked.push(what);
        if (navigated) {
          bProblems.push(`النقر على «${what}» أعاد تحميل الصفحة أو انتقل منها`);
          await page.goto(url, { waitUntil: 'load' });
          await ready(page);
          await stubPrint();
          continue;
        }
        if (errors.length > errBefore) bProblems.push(`النقر على «${what}» سبّب خطأ: ${short(errors[errors.length - 1].text, 160)}`);
        if (info.action === 'print') {
          const after = await page.evaluate(() => window.__apaPrintCalls);
          if (!(after > printsBefore)) bProblems.push('زر الطباعة [data-action=print] لم يستدعِ window.print()');
        }
      }
      if (!clicked.includes('print')) bProblems.push('لم يُعثر على زر طباعة ظاهر [data-action=print]');
      add('buttons', bProblems.length === 0 && clicked.length > 0,
        bProblems.length ? bProblems.join('؛ ') : `نُقر ${clicked.length} زرًا/إجراءً دون أخطاء، وزر الطباعة يستدعي window.print()`, { clicked });
    });
    await page.close().catch(() => {});

    // ---- print --------------------------------------------------------------------------------------------
    await section(['print'], async () => {
      const pp = await open(desktop, 'print');
      await pp.emulateMedia({ media: 'print' });
      await pp.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
      await settle(pp);
      const ph = await pp.evaluate((ids) => {
        const els = Array.prototype.slice.call(document.querySelectorAll('[data-print="hide"]'));
        const shownOnPaper = (el) => {
          if (!el) return false;
          for (let p = el; p && p !== document.documentElement; p = p.parentElement) {
            const cs = getComputedStyle(p);
            if (cs.display === 'none' || cs.visibility === 'hidden') return false;
          }
          const r = el.getBoundingClientRect();
          return r.width > 0 && r.height > 0;
        };
        const cards = Array.prototype.filter.call(document.querySelectorAll('[data-role="kpi-card"]'), shownOnPaper).length;
        return {
          n: els.length,
          shown: els.filter((e) => getComputedStyle(e).display !== 'none').map((e) => e.tagName.toLowerCase()
            + (e.getAttribute('data-role') ? `[data-role=${e.getAttribute('data-role')}]` : '')
            + (e.getAttribute('data-action') ? `[data-action=${e.getAttribute('data-action')}]` : '')),
          missing: ids.filter((id) => document.getElementById(id) && !shownOnPaper(document.getElementById(id))),
          absent: ids.filter((id) => !document.getElementById(id)),
          cards, cardsTotal: document.querySelectorAll('[data-role="kpi-card"]').length,
          // brief §6/§11: on paper a yield card still shows its scope, basis, valuation date and included count
          ratio: (() => {
            const bad = [];
            let all = null;
            try { all = window.APA.computeAll(window.APA_UI.getFilters()); } catch (e) { return ['computeAll: ' + String(e && e.message || e)]; }
            const norm = (x) => String(x || '').replace(/\s+/g, ' ').trim();
            document.querySelectorAll('[data-role="kpi-card"][data-kind="ratio"]').forEach((card) => {
              const id = card.getAttribute('data-metric-id');
              const m = all && all.metrics[id];
              if (!m || typeof m.value !== 'number' || !shownOnPaper(card)) return;
              const txt = norm(card.innerText);
              [['scope_text', 'النطاق'], ['basis', 'الأساس'], ['valuation_date', 'تاريخ التقييم']].forEach(([k, lab]) => {
                if (m[k] && !txt.includes(norm(m[k]))) bad.push(`${id}: ${lab} «${norm(m[k]).slice(0, 50)}» غير ظاهر على الورق`);
              });
              if (!/مشمول/.test(txt)) bad.push(`${id}: عدد المشمول والمستبعد غير ظاهر على الورق`);
            });
            return bad;
          })()
        };
      }, PRINT_SECTIONS);
      const pdfPath = path.join(screensDir, 'print.pdf');
      if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
      await pp.pdf({ path: pdfPath, format: 'A4', printBackground: true });
      const size = fs.existsSync(pdfPath) ? fs.statSync(pdfPath).size : 0;
      let pages = null;
      try { pages = (fs.readFileSync(pdfPath, 'latin1').match(/\/Type\s*\/Page(?!s)/g) || []).length || null; } catch (e) { /* size check below */ }
      if (size > 0) report.screenshots.push('screens/print.pdf');
      const prob = [];
      if (!ph.n) prob.push('لا توجد عناصر [data-print=hide] (المرشحات والأزرار يجب أن تُخفى عند الطباعة)');
      if (ph.shown.length) prob.push(`${ph.shown.length} عنصر [data-print=hide] يظهر عند الطباعة: ` + short(ph.shown.slice(0, 5)));
      if (ph.missing.length) prob.push('أقسام العرض التنفيذي لا تظهر عند الطباعة: ' + ph.missing.map((x) => '#' + x).join('، '));
      if (ph.absent.length) prob.push('أقسام العرض التنفيذي غير موجودة: ' + ph.absent.map((x) => '#' + x).join('، '));
      if (ph.cardsTotal && ph.cards < ph.cardsTotal) prob.push(`${nAr(ph.cardsTotal - ph.cards, 'بطاقة مؤشر واحدة', 'بطاقتا مؤشر', 'بطاقات مؤشرات', 'بطاقة مؤشر')} لا تظهر عند الطباعة`);
      if (!(size > 1000)) prob.push('ملف PDF للطباعة لم يُنشأ أو فارغ');
      if (ph.ratio.length) prob.push('بطاقة العائد على الورق: ' + ph.ratio.slice(0, 4).join('؛ '));
      if (pages && pages > MAX_PRINT_PAGES) prob.push(`الطباعة ${nAr(pages, 'صفحة واحدة', 'صفحتان', 'صفحات', 'صفحة')} A4 والحد ${MAX_PRINT_PAGES} (dashboard-spec §8)؛ اختصر نصوص العرض التنفيذي`);
      add('print', prob.length === 0, prob.length ? prob.join('؛ ')
        : `عناصر [data-print=hide] (${ph.n}) مخفية عند الطباعة، وأقسام العرض التنفيذي (${PRINT_SECTIONS.length}) وبطاقاته (${ph.cards}) ظاهرة، ونطاق بطاقات العائد وأساسها وتاريخ تقييمها على الورق، ` +
          `وأُنشئ screens/print.pdf (A4، ${pages ? nAr(pages, 'صفحة واحدة', 'صفحتان', 'صفحات', 'صفحة') + '، ' : ''}${Math.round(size / 1024)} ك.ب)`, pages ? { pages } : undefined);
      await pp.close();
    });
    await desktop.close().catch(() => {});

    // ---- mobile: 360 and 390 px, default view and with the details + an asset detail open -------------------
    await section(['mobile_overflow'], async () => {
      const measure = (mp) => mp.evaluate(() => {
        // Mobile Chromium zooms out to fit over-wide content, which inflates innerWidth to the content
        // width; the layout/visual viewport keeps the device width, so take the smallest of the three.
        const vw = Math.min(window.innerWidth, document.documentElement.clientWidth || Infinity,
          (window.visualViewport && window.visualViewport.width) || Infinity);
        const sw = document.documentElement.scrollWidth;
        const offenders = [];
        if (sw > vw + 1) {
          const clipped = (el) => {
            for (let p = el.parentElement; p && p !== document.body; p = p.parentElement) {
              const ox = getComputedStyle(p).overflowX;
              if (ox === 'auto' || ox === 'scroll' || ox === 'hidden' || ox === 'clip') return true;
            }
            return false;
          };
          document.querySelectorAll('body *').forEach((el) => {
            if (offenders.length >= 5) return;
            const r = el.getBoundingClientRect();
            if ((r.right > vw + 1 || r.left < -1) && r.width > 0 && !clipped(el)) {
              const cls = typeof el.className === 'string' && el.className.trim() ? '.' + el.className.trim().split(/\s+/).slice(0, 2).join('.') : '';
              offenders.push(el.tagName.toLowerCase() + (el.id ? '#' + el.id : '') + cls + ` (${Math.round(r.left)}..${Math.round(r.right)})`);
            }
          });
        }
        return { vw: Math.round(vw), sw, offenders };
      });
      const results = [];
      const mProblems = [];
      for (const w of MOBILE_WIDTHS) {
        const base = w === 390 ? Object.assign({}, pw.devices['iPhone 13']) : { viewport: { width: w, height: 780 }, deviceScaleFactor: 2, isMobile: true, hasTouch: true };
        delete base.defaultBrowserType;
        const mobile = await browser.newContext(Object.assign(base, { locale: 'ar-SA', timezoneId: 'Asia/Riyadh', bypassCSP: true }));
        await mobile.addInitScript({ content: PAGE_HELPERS });
        try {
          const mp = await open(mobile, `mobile${w}`);
          const d0 = await measure(mp);
          results.push({ width: w, state: 'default', ...d0 });
          if (w === 390) {
            await mp.screenshot({ path: path.join(screensDir, 'mobile.png') });
            report.screenshots.push('screens/mobile.png');
            try {
              await mp.screenshot({ path: path.join(screensDir, 'mobile_full.png'), fullPage: true, scale: 'css' });
              report.screenshots.push('screens/mobile_full.png');
            } catch (e) { /* very long pages: the viewport shot is enough */ }
          } else {
            await mp.screenshot({ path: path.join(screensDir, `mobile_${w}.png`) });
            report.screenshots.push(`screens/mobile_${w}.png`);
          }
          // every <details> open and the first asset's detail panel open
          await mp.evaluate(() => { document.querySelectorAll('details').forEach((d) => { d.open = true; }); });
          const row = mp.locator('table[data-role="asset-table"] tr[data-role="asset-row"]').first();
          let opened = false;
          if (await row.count()) {
            await row.scrollIntoViewIfNeeded().catch(() => {});
            await row.click().catch(async () => { await row.evaluate((el) => el.click()); });
            opened = true;
          }
          await waitStable(mp);
          const d1 = await measure(mp);
          results.push({ width: w, state: opened ? 'details_and_asset_detail_open' : 'details_open', ...d1 });
          if (w === MOBILE_WIDTHS[0] && opened) {
            try {
              // what the phone shows when a row is opened: the top of the detail panel, one screen tall
              await mp.evaluate(() => { const d = document.querySelector('[data-role="asset-detail"]'); if (d) d.scrollIntoView({ block: 'start' }); });
              await settle(mp);
              await mp.screenshot({ path: path.join(screensDir, 'mobile_detail.png') });
              report.screenshots.push('screens/mobile_detail.png');
            } catch (e) { /* the overflow numbers above are the check */ }
          }
        } finally {
          await mobile.close().catch(() => {});
        }
      }
      for (const r of results) {
        if (r.sw > r.vw + 1) mProblems.push(`${r.width}px (${r.state === 'default' ? 'العرض الافتراضي' : 'مع فتح التفاصيل وتفاصيل أصل'}): scrollWidth=${r.sw} > ${r.vw}؛ ` +
          `عناصر تتجاوز العرض: ${r.offenders.join('، ') || 'غير محددة'}`);
      }
      add('mobile_overflow', mProblems.length === 0, mProblems.length ? 'تمرير أفقي على الجوال: ' + mProblems.join('؛ ')
        : `لا تمرير أفقي على ${MOBILE_WIDTHS.join(' و')} بكسل في العرض الافتراضي ومع فتح كل التفاصيل وتفاصيل أصل؛ اللقطات: screens/mobile.png، screens/mobile_${MOBILE_WIDTHS[0]}.png، screens/mobile_detail.png`,
        results);
    });

    // ---- global -------------------------------------------------------------------------------------------
    const ext = [...new Set(external.map((x) => x.url))];
    add('offline', ext.length === 0, ext.length ? 'الصفحة تطلب موارد من الشبكة (يجب أن تعمل دون اتصال): ' + short(ext.slice(0, 5), 400)
      : 'لا طلبات شبكة خارجية؛ الصفحة مكتفية ذاتيًا');
    add('console_errors', errors.length === 0, errors.length
      ? `${errors.length} خطأ في وحدة التحكم/الصفحة: ` + short(errors.slice(0, 5).map((x) => `[${x.page}] ${x.text}`).join(' | '), 700)
      : 'لا أخطاء في وحدة التحكم ولا أخطاء صفحة (سطح المكتب، الطباعة، الجوال)');
  } catch (e) {
    add('runner', false, 'توقف الفحص بسبب خطأ غير متوقع: ' + short(String((e && e.stack) || e), 600));
  } finally {
    await browser.close().catch(() => {});
  }
  return finish();
}

main().then((code) => process.exit(code), (e) => {
  console.error('خطأ غير متوقع في check_dashboard.mjs: ' + ((e && e.stack) || e));
  process.exit(1);
});
