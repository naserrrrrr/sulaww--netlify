#!/usr/bin/env node
/**
 * check_dashboard.mjs — headless-browser test of a built dashboard.html (data-contract.md §0, §8).
 *
 * Usage:
 *   node scripts/check_dashboard.mjs --html RUN/dashboard.html --out RUN [--timeout 30000]
 * Writes RUN/08_dashboard_check.json and RUN/screens/{desktop.png, mobile.png, mobile_full.png, print.pdf}.
 * Exit code: 0 = every check passed, 1 = a check failed (or the browser could not run), 2 = bad arguments.
 *
 * Checks (DOM contract shared with assets/dashboard_template.html):
 *   load, html_lang_dir, api, structure, details_sections, parity (every parity case of
 *   05_metrics.json recomputed in the page by APA.computeAll), cards, charts, filters, period_filter,
 *   table_search, table_sort, table_detail, buttons, no_raw_tokens, privacy, desktop_screenshot,
 *   mobile_overflow, print, offline, console_errors.
 * Playwright is loaded from the local project, NODE_PATH or the global npm root.
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

const PARITY_REL_TOL = 1e-9;
const MAX_FILTER_VALUES = 3;
const MAX_BUTTONS = 150;
const MAX_SORT_HEADERS = 6;
const DETAIL_SECTIONS = ['calculations', 'issues', 'exclusions', 'assumptions', 'sources', 'workbook', 'questions',
  'scenarios', 'versions', 'methodology', 'records'];
const MAX_CARDS = 8;

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------
function usage(msg) {
  if (msg) console.error('خطأ: ' + msg);
  console.error('الاستخدام: node scripts/check_dashboard.mjs --html RUN/dashboard.html --out RUN [--timeout 30000]');
  process.exit(2);
}

function parseArgs(argv) {
  const out = { timeout: 30000 };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const eq = a.indexOf('=');
    const key = eq > 0 ? a.slice(0, eq) : a;
    const val = () => (eq > 0 ? a.slice(eq + 1) : argv[++i]);
    if (key === '--html') out.html = val();
    else if (key === '--out') out.out = val();
    else if (key === '--timeout') out.timeout = Number(val());
    else if (key === '-h' || key === '--help') usage();
    else usage('وسيط غير معروف: ' + a);
  }
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
      if (!vals.length) out.push({ card: ci, cardId: card.getAttribute('data-metric-id'), missing: true });
      vals.forEach(function (el) {
        var holder = el.closest('[data-metric-id]') || card;
        out.push({
          card: ci, cardId: card.getAttribute('data-metric-id'), metricId: holder.getAttribute('data-metric-id'),
          dataValue: el.getAttribute('data-value'), dataStatus: el.getAttribute('data-status'),
          text: (el.innerText || el.textContent || '').trim(), visible: visible(el)
        });
      });
    });
    return out;
  }
  function charts() {
    return Array.prototype.map.call(document.querySelectorAll('[data-role="chart"]'), function (c) {
      var empty = c.querySelector('[data-role="chart-empty"]');
      return { id: c.getAttribute('data-chart-id'), svg: !!c.querySelector('svg'),
        empty: !!empty, emptyText: empty ? (empty.textContent || '').trim() : null };
    });
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
  function snapshot() {
    var out = { error: null };
    try {
      out.filters = window.APA_UI.getFilters();
      var all = window.APA.computeAll(out.filters);
      out.expected = compact(all);
      out.expectedRows = window.APA.filterRecords(out.filters).length;
    } catch (e) { out.error = String(e && e.message || e); }
    out.rows = rowIds();
    out.cards = cards();
    out.charts = charts();
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
    snapshot: snapshot, signature: signature, table: table, compact: compact };
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
    }
    if (/NaN|undefined|Infinity|\[object/.test(c.text)) problems.push(`${id}: نص غير صالح في البطاقة «${c.text}»`);
  }
  return problems;
}

function compareCharts(snap) {
  const problems = [];
  for (const ch of snap.charts) {
    if (!ch.svg && !ch.empty) problems.push(`الرسم ${ch.id}: لا يحتوي <svg> ولا [data-role=chart-empty]`);
    else if (!ch.svg && ch.empty && !ch.emptyText) problems.push(`الرسم ${ch.id}: حالة فارغة دون سبب مكتوب`);
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
async function main() {
  const args = parseArgs(process.argv.slice(2));
  const htmlPath = path.resolve(args.html);
  const outDir = path.resolve(args.out);
  const screensDir = path.join(outDir, 'screens');
  fs.mkdirSync(screensDir, { recursive: true });
  const report = {
    contract_version: '1.0', html: htmlPath, checked_at: new Date().toISOString(), passed: false,
    summary: null, checks: [], screenshots: [], console_errors: [], external_requests: []
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

    // ---- desktop screenshot of the untouched default view ---------------------------------------
    await section(['desktop_screenshot'], async () => {
      await page.screenshot({ path: path.join(screensDir, 'desktop.png'), fullPage: true });
      report.screenshots.push('screens/desktop.png');
      add('desktop_screenshot', true, 'لقطة سطح المكتب 1440×900 (الصفحة كاملة): screens/desktop.png');
    });

    // ---- html lang / dir, API, structure -----------------------------------------------------------
    let st = { reset: 0, table: 0, search: 0, sortable: 0, sections: [] };
    await section(['html_lang_dir', 'api', 'structure', 'details_sections'], async () => {
      const ld = await page.evaluate(() => ({ lang: document.documentElement.getAttribute('lang'), dir: document.documentElement.getAttribute('dir') }));
      add('html_lang_dir', ld.lang === 'ar' && ld.dir === 'rtl', `lang="${ld.lang}" dir="${ld.dir}" (المطلوب lang="ar" dir="rtl")`);

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
    await section(['cards', 'charts'], async () => {
      def = await snap(page);
      const cardIds = await page.evaluate(() => ((window.APA.data.config && window.APA.data.config.metrics) || [])
        .filter((m) => m && m.card).map((m) => m.id));
      const cardProblems = compareCards(def);
      const shown = new Set(def.cards.map((c) => c.metricId));
      const notShown = cardIds.filter((id) => !shown.has(id));
      if (notShown.length) cardProblems.push('مؤشرات card:true غير معروضة في البطاقات: ' + notShown.join('، '));
      add('cards', cardProblems.length === 0 && def.cards.length > 0,
        cardProblems.length ? cardProblems.join('؛ ') : `${def.cards.length} قيمة بطاقة تطابق computeAll للمرشحات الافتراضية`,
        { filters: def.filters, cards: def.cards.map((c) => ({ id: c.metricId, value: c.dataValue, status: c.dataStatus, text: c.text })) });
      const chartProblems = compareCharts(def);
      add('charts', chartProblems.length === 0 && def.charts.length > 0,
        def.charts.length === 0 ? 'لا توجد رسوم [data-role=chart]'
          : chartProblems.length ? chartProblems.join('؛ ') : `${def.charts.length} رسمًا: كل رسم يحتوي SVG أو حالة فارغة مسببة`);
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
            : `جُرّب ${tried.length} خيارًا في ${dimSelects} مرشحات: الصفوف والبطاقات والرسوم ونص النطاق متسقة، وإعادة الضبط تستعيد الافتراضي`,
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
        await header(i).click();
        await waitStable(page);
        const s1 = await page.evaluate((idx) => {
          const ths = document.querySelectorAll('table[data-role="asset-table"] th[data-sort]');
          const ci = ths[idx] ? ths[idx].cellIndex : -1;
          const rows = window.__apaCheck.rows();
          return { ids: rows.map((r) => r.getAttribute('data-record-id')),
            texts: rows.map((r) => (r.cells[ci] ? (r.cells[ci].textContent || '').trim() : null)) };
        }, i);
        await header(i).click();
        await waitStable(page);
        const ids2 = await page.evaluate(() => window.__apaCheck.rowIds());
        // Rows without a value ("—", "غير متاح", ...) may legitimately stay last in both directions:
        // only the rows holding a value must reverse, and only when they hold >= 2 distinct values.
        const textOf = new Map(s1.ids.map((id, k) => [id, s1.texts[k]]));
        const hasValue = (id) => { const t = textOf.get(id); return t !== null && t !== undefined && !isMissingText(t); };
        const valued1 = s1.ids.filter(hasValue), valued2 = ids2.filter(hasValue);
        const distinct = new Set(valued1.map((id) => textOf.get(id))).size;
        const changed = !sameJSON(valued1, valued2);
        sortInfo.push({ key, rows: ids2.length, rows_with_value: valued1.length, distinct_values: distinct, changed });
        if (!sameJSON([...s1.ids].sort(), [...idsBefore].sort()) || !sameJSON([...ids2].sort(), [...idsBefore].sort())) {
          sProblems.push(`الفرز حسب ${key} غيّر مجموعة الصفوف`);
        } else if (distinct >= 2 && !changed) {
          sProblems.push(`النقر مرتين على عمود ${key} لم يعكس ترتيب الصفوف رغم اختلاف القيم`);
        }
      }
      if (!nHeaders) sProblems.push('لا توجد أعمدة قابلة للفرز th[data-sort]');
      add('table_sort', sProblems.length === 0, sProblems.length ? sProblems.join('؛ ') : `جُرّب الفرز (نقرتان) على ${sortInfo.length} عمودًا`, sortInfo);
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
            if (key === 'raw' && parent && typeof parent.v === 'number' && Number(x.replace(/[,\s٬]/g, '')) === parent.v) return;
            out.push(x);
          } else if (Array.isArray(x)) x.forEach((y) => walk(y, key, parent));
          else if (x && typeof x === 'object') Object.keys(x).forEach((k) => walk(x[k], k, x));
        };
        walk(window.APA.data, null, null);
        return out;
      });
      const BAD = /\bNaN\b|\bundefined\b|\[object Object\]|\bInfinity\b|\{\{\s*m:/;
      const badHits = texts.filter((t) => BAD.test(t)).map((t) => short(t.trim(), 120));
      add('no_raw_tokens', badHits.length === 0, badHits.length
        ? 'نصوص غير صالحة في الصفحة (NaN/undefined/[object Object]/Infinity أو رموز {{m:}} غير مُعالجة): ' + short(badHits.slice(0, 5), 500)
        : 'لا توجد قيم خام غير صالحة في نص الصفحة', badHits.length ? badHits.slice(0, 20) : undefined);
      const PHONE = /(?<!\d)(?:\+?966|00966|0)?5\d{8}(?!\d)/;
      const EMAIL = /[\w.+-]+@[\w-]+\.[\w.-]+/;
      const piiHits = [];
      const scan = (arr, where) => {
        for (const t0 of arr) {
          const t = toAsciiDigits(t0);
          const collapsed = t.replace(/(\d)[\s\u00A0-](?=\d)/g, '$1');
          if (PHONE.test(t) || PHONE.test(collapsed)) piiHits.push({ where, kind: 'phone' });
          else if (EMAIL.test(t)) piiHits.push({ where, kind: 'email' });
        }
      };
      scan(texts, 'page_text');
      scan(dataStrings, 'embedded_data');
      const piiSummary = piiHits.reduce((a, h) => { const k = h.where + ':' + h.kind; a[k] = (a[k] || 0) + 1; return a; }, {});
      add('privacy', piiHits.length === 0,
        piiHits.length ? 'وُجدت أنماط بيانات شخصية (لا تُطبع القيم): ' + Object.entries(piiSummary).map(([k, n]) => `${k}×${n}`).join('، ')
          : `لا أرقام جوال سعودية ولا بريد إلكتروني في نص الصفحة (${texts.length} نصًا) ولا في البيانات المضمّنة (${dataStrings.length} نصًا)`,
        piiHits.length ? piiSummary : undefined);
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
      await settle(pp);
      const ph = await pp.evaluate(() => {
        const els = Array.prototype.slice.call(document.querySelectorAll('[data-print="hide"]'));
        return { n: els.length, shown: els.filter((e) => getComputedStyle(e).display !== 'none').map((e) => e.tagName.toLowerCase()
          + (e.getAttribute('data-role') ? `[data-role=${e.getAttribute('data-role')}]` : '')
          + (e.getAttribute('data-action') ? `[data-action=${e.getAttribute('data-action')}]` : '')) };
      });
      const pdfPath = path.join(screensDir, 'print.pdf');
      if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
      await pp.pdf({ path: pdfPath, format: 'A4', printBackground: true });
      const size = fs.existsSync(pdfPath) ? fs.statSync(pdfPath).size : 0;
      if (size > 0) report.screenshots.push('screens/print.pdf');
      const prob = [];
      if (!ph.n) prob.push('لا توجد عناصر [data-print=hide] (المرشحات والأزرار يجب أن تُخفى عند الطباعة)');
      if (ph.shown.length) prob.push(`${ph.shown.length} عنصر [data-print=hide] يظهر عند الطباعة: ` + short(ph.shown.slice(0, 5)));
      if (!(size > 1000)) prob.push('ملف PDF للطباعة لم يُنشأ أو فارغ');
      add('print', prob.length === 0, prob.length ? prob.join('؛ ')
        : `${ph.n} عنصر [data-print=hide] مخفي عند الطباعة، وأُنشئ screens/print.pdf (A4، ${Math.round(size / 1024)} ك.ب)`);
      await pp.close();
    });
    await desktop.close().catch(() => {});

    // ---- mobile -------------------------------------------------------------------------------------------
    await section(['mobile_overflow'], async () => {
      const dev = Object.assign({}, pw.devices['iPhone 13']);
      delete dev.defaultBrowserType;
      const mobile = await browser.newContext(Object.assign(dev, { locale: 'ar-SA', timezoneId: 'Asia/Riyadh', bypassCSP: true }));
      await mobile.addInitScript({ content: PAGE_HELPERS });
      try {
        const mp = await open(mobile, 'mobile');
        const ov = await mp.evaluate(() => {
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
          return { vw: Math.round(vw), sw, offenders, innerWidth: window.innerWidth };
        });
        await mp.screenshot({ path: path.join(screensDir, 'mobile.png') });
        report.screenshots.push('screens/mobile.png');
        try {
          await mp.screenshot({ path: path.join(screensDir, 'mobile_full.png'), fullPage: true, scale: 'css' });
          report.screenshots.push('screens/mobile_full.png');
        } catch (e) { /* very long pages: the viewport shot is enough */ }
        add('mobile_overflow', ov.sw <= ov.vw + 1,
          ov.sw <= ov.vw + 1 ? `لا تمرير أفقي على iPhone 13 (عرض الصفحة ${ov.sw} ≤ ${ov.vw})، اللقطات: screens/mobile.png وscreens/mobile_full.png`
            : `تمرير أفقي على الجوال: scrollWidth=${ov.sw} > عرض الشاشة الفعلي ${ov.vw}؛ عناصر تتجاوز العرض: ${ov.offenders.join('، ')}`, ov);
      } finally {
        await mobile.close().catch(() => {});
      }
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
