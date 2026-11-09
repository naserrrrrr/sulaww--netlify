#!/usr/bin/env python3
"""Board-review fixes of the dashboard, on a FICTIONAL one-asset run (no real workbook value).

    python3 -I tests/test_dashboard_board.py [--work DIR] [--skip-browser]      # exit 0 = all pass, 1 = a failure

Build (no browser):
  - «تاريخ إعداد الدراسة»: a study date taken from the workbook's last-modified property is not the study date
    (meta.study_date = null, file_modified kept as context); the confidentiality marking is on by default;
    the Arabic web font is embedded base64 in @font-face; the printed next-step hint names check_dashboard.mjs by an
    absolute, quoted path; per-record facts (rec.*) are embedded only when the narrative cites them.
Page (Playwright; skipped with --skip-browser or APA_SKIP_BROWSER=1):
  - C1: a metric spec with review shows the amber «يحتاج تحقق» chip and its reason, never the green «مكتمل» chip,
    under every period; the calculations log names it too.
  - C3: the asset-table data-note chip counts the same list as the detail panel (03 issues on the record, issues naming
    it, 06 data_issues_extra naming it, a 01 anomaly attached to it).
  - C4: the top actions follow the recommendations' priority and are numbered by it.
  - one asset with values: one sentence replaces the distribution/concentration/comparison placeholders; the trend
    is full width; no empty placeholder box prints.
  - the location column labels a district shown in place of an unknown location («الحي: …»).
  - consecutive exclusions sharing one reason are grouped under one reason line; every row stays.
  - ≤ 600 px: asset rows are cards (value, revenue, yield and status inside the screen), no header band.
  - print: identity captions print whole (no «…»), confidentiality in the header and the page footer.
  - check_dashboard.mjs passes on the run (cards/review, print, mobile_table, priority_order, table_detail …).
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

sys.dont_write_bytecode = True
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
S = os.path.join(SKILL, "scripts")
PY = [sys.executable, "-I"]
FAILS, PASSES = [], []

REVIEW_REASON = "إيراد السنة الأخيرة غير محسوم لوجود رقم بديل في بيان تجريبي آخر"
SHARED_REASON = "ورقة تجريبية مخفية لم يتحقق انتماء بنودها إلى المحفظة فلا تدخل أي مؤشر"
MODIFIED = "2031-03-14"


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {str(info)[:500]}"))


def run(cmd, **kw):
    p = subprocess.run(cmd, capture_output=True, text=True, **kw)
    return p.returncode, p.stdout + p.stderr


def dump(p, o):
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(o, fh, ensure_ascii=False, indent=1)


def load(p):
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def tv(v, ref):
    return {"v": v, "s": "value", "raw": str(v), "ref": ref}


FIELD_META = {
    "rev.2030.demo": {"label": "إيراد تجريبي 2030", "group": "revenue", "period": "2030", "vat": "excl", "basis": "recorded",
                      "cash_basis": "unknown", "completeness": "unknown", "unit": "SAR"},
    "rev.2029.demo": {"label": "إيراد تجريبي 2029", "group": "revenue", "period": "2029", "vat": "excl", "basis": "recorded",
                      "cash_basis": "unknown", "completeness": "unknown", "unit": "SAR"},
    "val.2029.demo": {"label": "تقييم تجريبي 2029", "group": "valuation", "year": "2029", "valuation_basis": "unknown",
                      "valuation_date": "2029 (اليوم غير محدد)", "component": "total", "unit": "SAR"},
}


def make_run(d):
    """One in-scope asset (district known, location unknown) and five isolated lines, four sharing one reason."""
    os.makedirs(d, exist_ok=True)
    asset = {"id": "T-1", "name": "مبنى تجريبي أول", "record_type": "asset", "scope": "in_scope",
             "dims": {"use": "investment", "income_nature": "rent", "asset_type": "عمارة", "location": None,
                      "district": "حي تجريبي"},
             "source": {"source_id": "REG", "sheet": "سجل تجريبي", "row": 3, "ref_name": "سجل تجريبي!B3"},
             "fields": {"rev.2030.demo": tv(120000, "سجل تجريبي!C3"), "rev.2029.demo": tv(100000, "سجل تجريبي!D3"),
                        "val.2029.demo": tv(2400000, "سجل تجريبي!E3")},
             "issue_ids": ["DQ-901"]}
    isolated = []
    for k in range(1, 6):
        isolated.append({"id": f"X-{k}", "name": f"بند معزول تجريبي {k}", "record_type": "asset", "scope": "isolated",
                         "scope_reason": SHARED_REASON if k <= 4 else "سبب تجريبي منفرد لبند واحد",
                         "dims": {}, "source": {"source_id": "HID", "sheet": "ورقة تجريبية", "row": k + 1,
                                                "ref_name": f"ورقة تجريبية!B{k + 1}"},
                         "fields": {"rev.2030.demo": tv(1000 * k, f"ورقة تجريبية!C{k + 1}")}})
    issues = [
        {"id": "DQ-901", "type": "date_note", "severity": "low", "where": "سجل تجريبي!F3", "record_ids": ["T-1"],
         "detail": "ملاحظة تجريبية أولى على السجل", "impact": "أثر تجريبي", "treatment": "معالجة تجريبية"},
        {"id": "DQ-902", "type": "date_note", "severity": "medium", "where": "سجل تجريبي!G3", "record_ids": ["T-1"],
         "detail": "ملاحظة تجريبية ثانية تسمّي السجل", "impact": "أثر تجريبي", "treatment": "معالجة تجريبية"}]
    dump(os.path.join(d, "03_records.json"), {
        "contract_version": "1.0", "source": {"name": "fictional_register.xlsx", "sha256": "0" * 64},
        "field_meta": FIELD_META, "records": [asset] + isolated, "excluded_rows": [], "issues": issues,
        "stats": {"records": 6}})
    dump(os.path.join(d, "01_workbook_map.json"), {
        "file": {"name": "fictional_register.xlsx", "properties": {"modified": MODIFIED + "T09:30:00"}},
        "summary": {"sheets": 2, "hidden_sheets": 1}, "sheets": [{"name": "سجل تجريبي", "state": "visible"},
                                                                {"name": "ورقة تجريبية", "state": "hidden"}],
        "anomalies": [{"id": "WB-901", "type": "note", "severity": "low", "where": "سجل تجريبي!H3",
                       "detail": "ملاحظة مصنف تجريبية مربوطة بالسجل", "record_ids": ["T-1"]}]})
    dump(os.path.join(d, "02_mapping.json"), {
        "entity": {"name": None, "evidence": "لا يذكر الملف التجريبي اسم الجهة"},
        "data_period": {"label": "2030", "evidence": "سجل تجريبي!A1"},
        "study_date": {"value": MODIFIED, "basis": "آخر تعديل للملف (خصائص المصنف)", "evidence": "file.properties.modified"},
        "analysis_date": "2031-05-01", "currency": "SAR", "valuation_dates": [], "sheets": []})


def narrative():
    rec = lambda rid, pr, title: {  # noqa: E731
        "id": rid, "priority": pr, "title": title, "target": "T-1", "gap_ids": ["G1"], "problem": "مشكلة تجريبية",
        "first_action": "إجراء تجريبي", "owner_role": "مدير تجريبي", "duration": "30 يومًا",
        "expected_impact": "أثر تجريبي متوقع", "evidence": ["m:recorded_value"]}
    return {
        "contract_version": "1.0",
        "executive_summary": "خلاصة تجريبية قصيرة لأصل واحد ضمن النطاق.",
        "gaps": [{"id": "G1", "title": "فجوة تجريبية", "scope": "T-1", "severity": "medium", "is_proven_problem": False,
                  "evidence": [{"ref": "m:recorded_value", "kind": "derived"}], "impact": {"text": "أثر تجريبي"}}],
        # the narrative lists R1 (priority 2) before R2 (priority 1): the page must follow the priority (C4)
        "recommendations": [rec("R1", 2, "توصية تجريبية ثانية في الأولوية"), rec("R2", 1, "توصية تجريبية أولى في الأولوية")],
        "top_gaps": ["G1"], "top_actions": ["R1", "R2"],
        "plan": {"d30": [{"action": "بند خطة تجريبي", "owner_role": "مدير تجريبي", "rec_id": "R2"}],
                 "d60": [{"action": "بند خطة تجريبي ثان", "owner_role": "مدير تجريبي", "rec_id": "R1"}], "d90": []},
        "data_issues_extra": [{"id": "DQ-X9", "where": "سجل تجريبي!B3", "record_ids": ["T-1"],
                               "detail": "ملاحظة محلل تجريبية على السجل", "impact": "أثر", "treatment": "معالجة"}]}


DOM_PROBE = r"""
import { createRequire } from 'module';
import { execSync } from 'child_process';
import path from 'path';
let pw = null;
for (const root of [process.cwd() + '/x.js', path.join(execSync('npm root -g', {encoding: 'utf8'}).trim(), 'x.js')]) {
  try { pw = createRequire(root)('playwright'); break; } catch (e) { /* next */ }
}
if (!pw) { console.log(JSON.stringify({error: 'playwright'})); process.exit(0); }
const url = 'file://' + process.argv[2];
const b = await pw.chromium.launch();
const errors = [];
const d = await b.newContext({ viewport: { width: 1440, height: 900 } });
const p = await d.newPage();
p.on('pageerror', (e) => errors.push(String(e)));
await p.goto(url);
await p.waitForFunction(() => !!window.APA_UI);
await p.waitForTimeout(300);
const out = { errors };
const cardState = () => p.evaluate(() => {
  const card = document.querySelector('[data-role="kpi-card"][data-metric-id="' + window.__reviewId + '"]');
  return card ? { review: card.getAttribute('data-review'), chip: !!card.querySelector('[data-role="review-chip"]'),
    ok: !!card.querySelector('.kpi-status .tone-ok'), text: card.innerText } : null;
});
out.reviewId = await p.evaluate(() => {
  const s = (window.APA.data.config.metrics || []).filter((m) => m.review && m.card)[0];
  window.__reviewId = s ? s.id : null; return window.__reviewId;
});
out.periods = [];
const periods = await p.evaluate(() => Array.from(document.querySelectorAll('select[data-filter="period"] option')).map((o) => o.value));
for (const v of periods) {
  await p.selectOption('select[data-filter="period"]', v);
  await p.waitForTimeout(150);
  out.periods.push(Object.assign({ period: v }, await cardState()));
}
await p.click('[data-action="reset-filters"]');
await p.waitForTimeout(150);
Object.assign(out, await p.evaluate(() => {
  const q = (s) => document.querySelector(s);
  const qa = (s) => Array.from(document.querySelectorAll(s));
  const calc = q('#d-calculations article.calc[data-metric-id="' + window.__reviewId + '"]');
  return {
    conf: (q('[data-role="confidentiality"]') || {}).textContent || null,
    confFooter: (q('[data-role="confidentiality-footer"]') || {}).textContent || null,
    pageStyle: (q('#page-footer-style') || {}).textContent || '',
    fontFaces: qa('style').map((s) => (s.textContent.match(/@font-face/g) || []).length).reduce((a, x) => a + x, 0),
    meta: qa('#meta-grid > div').map((x) => x.innerText.replace(/\s+/g, ' ')),
    calcReview: calc ? calc.textContent : null,   // inside a closed <details>: textContent, not innerText
    actions: qa('[data-role="top-action"]').map((a) => [a.getAttribute('data-rec-id'), (a.querySelector('.rank') || {}).textContent]),
    charts: qa('#chart-grid [data-role="chart"]').map((c) => [c.getAttribute('data-chart-id'), c.classList.contains('wide'), c.getAttribute('data-print')]),
    scopeNote: (q('[data-chart-id="asset-panels"]') || {}).innerText || null,
    loc: (q('tr[data-role="asset-row"] td[data-col="location"]') || {}).textContent || null,
    locHead: (qa('#asset-table thead th')[2] || {}).textContent || null,
    chip: (q('tr[data-role="asset-row"] td[data-col="flags"]') || {}).getAttribute ? q('tr[data-role="asset-row"] td[data-col="flags"]').getAttribute('data-issues') : null,
    exclGroups: qa('[data-role="exclusions-table"] tr[data-role="reason-group"]').map((r) => [r.getAttribute('data-count'), r.textContent]),
    exclRows: qa('[data-role="exclusions-table"] tbody tr:not([data-role="reason-group"])').length,
    kpiText: qa('[data-role="kpi-card"]').map((c) => c.innerText).join('\n'),
  };
}));
await p.click('tr[data-role="asset-row"]');
await p.waitForTimeout(150);
out.panel = await p.evaluate(() => { const s = document.querySelector('[data-role="record-issues"]'); return s ? [s.getAttribute('data-count'), s.innerText] : null; });
// print media at the A4 content width
await p.emulateMedia({ media: 'print' });
await p.setViewportSize({ width: 703, height: 1013 });
await p.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
await p.waitForTimeout(300);
out.print = await p.evaluate(() => ({
  captions: Array.from(document.querySelectorAll('#meta-grid dd')).map((dd) => dd.innerText.replace(/\s+/g, ' ')),
  shownCharts: Array.from(document.querySelectorAll('#chart-grid [data-role="chart"]')).filter((c) => getComputedStyle(c).display !== 'none').map((c) => c.getAttribute('data-chart-id')),
  confShown: (() => { const m = document.querySelector('[data-role="confidentiality"]'); return !!m && getComputedStyle(m).display !== 'none'; })(),
}));
// phone: asset rows as cards
const m = await b.newContext(Object.assign({}, pw.devices['iPhone 13'], { defaultBrowserType: undefined }));
const mp = await m.newPage();
mp.on('pageerror', (e) => errors.push(String(e)));
await mp.goto(url);
await mp.waitForFunction(() => !!window.APA_UI);
await mp.waitForTimeout(300);
out.mobile = await mp.evaluate(() => {
  const vw = document.documentElement.clientWidth;
  const t = document.querySelector('table[data-role="asset-table"]');
  const thead = t.querySelector('thead');
  const row = t.querySelector('tr[data-role="asset-row"]');
  row.scrollIntoView({ block: 'start' });
  const cells = {};
  ['value', 'revenue', 'yield', 'flags'].forEach((k) => { const r = row.querySelector('td[data-col="' + k + '"]').getBoundingClientRect(); cells[k] = [Math.round(r.left), Math.round(r.right), Math.round(r.height)]; });
  return { vw, theadShown: getComputedStyle(thead).display !== 'none', wrapSW: t.parentElement.scrollWidth, wrapCW: t.parentElement.clientWidth, cells,
    sortShown: getComputedStyle(document.querySelector('.mobile-sort')).display !== 'none' };
});
console.log(JSON.stringify(out));
await b.close();
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", default=None)
    ap.add_argument("--skip-browser", action="store_true")
    a = ap.parse_args()
    skip_browser = a.skip_browser or os.environ.get("APA_SKIP_BROWSER") == "1"
    work = a.work or tempfile.mkdtemp(prefix="apa_board_")
    d = os.path.join(work, "run")
    shutil.rmtree(d, ignore_errors=True)
    print("[build] تشغيل وهمي بأصل واحد")
    make_run(d)
    rc, log = run(PY + [os.path.join(S, "compute_metrics.py"), "--init-config", "--records", os.path.join(d, "03_records.json"),
                        "--out", os.path.join(d, "04_metrics_config.json")])
    check("compute_metrics --init-config", rc == 0, log[-400:])
    cfg = load(os.path.join(d, "04_metrics_config.json"))
    card_ids = [m["id"] for m in cfg.get("metrics") or [] if m.get("card")]
    rid = "property_revenue" if "property_revenue" in card_ids else (card_ids[0] if card_ids else None)
    for m in cfg.get("metrics") or []:
        if m["id"] == rid:
            m["review"] = {"status": "needs_verification", "reason": REVIEW_REASON}
    dump(os.path.join(d, "04_metrics_config.json"), cfg)
    rc, log = run(PY + [os.path.join(S, "compute_metrics.py"), "--records", os.path.join(d, "03_records.json"),
                        "--config", os.path.join(d, "04_metrics_config.json"), "--out", os.path.join(d, "05_metrics.json")])
    check("compute_metrics", rc == 0, log[-400:])
    nar = narrative()
    dump(os.path.join(d, "06_narrative.json"), nar)
    dump(os.path.join(d, "07_validation.json"), {"passed": True, "blocking_count": 0, "warning_count": 0, "checks": [],
                                                 "generated_at": "2031-05-01T10:00:00"})
    rc, log = run(PY + [os.path.join(S, "build_dashboard.py"), "--run-dir", d], cwd=work)
    check("build_dashboard exits 0", rc == 0, log[-600:])
    if rc != 0:
        return finish()
    hint = [ln for ln in log.splitlines() if "check_dashboard.mjs" in ln]
    check("next-step hint: absolute quoted script and run paths",
          bool(hint) and f'"{os.path.join(S, "check_dashboard.mjs")}"' in hint[0] and f'"{os.path.abspath(d)}"' in hint[0], hint)
    an = load(os.path.join(d, "analysis.json"))
    meta = an.get("meta") or {}
    check("study date from file properties is not the study date",
          meta.get("study_date") is None and meta.get("study_date_stated") is False and meta.get("file_modified") == MODIFIED, meta)
    check("confidentiality marking on by default", meta.get("confidentiality") == "سري — للاستخدام الداخلي", meta.get("confidentiality"))
    with open(os.path.join(d, "dashboard.html"), encoding="utf-8") as fh:
        html = fh.read()
    check("Arabic web font embedded (base64 @font-face, no URL)",
          html.count("@font-face{font-family:'IBM Plex Sans Arabic'") == 6 and "data:font/woff2;base64," in html
          and "/*__APA_FONTS__*/" not in html, html.count("@font-face"))
    check("anomaly record_ids kept for the data-note list",
          any(x.get("record_ids") == ["T-1"] for x in (an.get("workbook_map_summary") or {}).get("anomalies") or []))
    if skip_browser:
        print("  SKIP  DOM checks (--skip-browser)")
        return finish()

    print("[page] الصفحة في المتصفح")
    probe = os.path.join(work, "board_probe.mjs")
    with open(probe, "w", encoding="utf-8") as fh:
        fh.write(DOM_PROBE)
    rc, log = run(["node", probe, os.path.join(d, "dashboard.html")], cwd=work)
    try:
        dom = json.loads(log.strip().splitlines()[-1])
    except (ValueError, IndexError):
        dom = {"error": log[-400:]}
    if dom.get("error"):
        check("DOM probe ran", False, dom)
        return finish()
    check("no page errors", not dom.get("errors"), dom.get("errors"))
    per = dom.get("periods") or []
    check("C1: review chip and reason, no green chip, in every period",
          dom.get("reviewId") == rid and len(per) >= 2 and all(x and x.get("review") == "needs_verification" and x.get("chip")
                                                              and not x.get("ok") and REVIEW_REASON in x.get("text", "") for x in per), per)
    check("C1: the calculations log names the review", REVIEW_REASON in (dom.get("calcReview") or ""), dom.get("calcReview"))
    check("C3: table chip = detail panel = 4 notes (03 ×2, 06 extra, 01 anomaly)",
          dom.get("chip") == "4" and (dom.get("panel") or [None])[0] == "4", (dom.get("chip"), dom.get("panel")))
    check("C4: top actions in priority order, numbered by priority",
          dom.get("actions") == [["R2", "1"], ["R1", "2"]], dom.get("actions"))
    charts = dom.get("charts") or []
    ids = [c[0] for c in charts]
    check("one asset: one sentence replaces the comparison panels; the trend is full width",
          ids[:1] == ["asset-panels"] and not {"concentration", "asset-comparison", "distribution"} & set(ids)
          and all(c[1] for c in charts if c[0] == "revenue-trend") and "مبنى تجريبي أول" in (dom.get("scopeNote") or ""), charts)
    check("location column labels the district", (dom.get("loc") or "").strip() == "الحي: حي تجريبي"
          and "الحي" in (dom.get("locHead") or ""), (dom.get("loc"), dom.get("locHead")))
    check("exclusions: four rows with one shared reason grouped under one line, all five rows kept",
          [g[0] for g in dom.get("exclGroups") or []] == ["4"] and SHARED_REASON in dom["exclGroups"][0][1]
          and dom.get("exclRows") == 5, (dom.get("exclGroups"), dom.get("exclRows")))
    kpi = dom.get("kpiText") or ""
    check("KPI count lines agree in Arabic: «لا سجلات مشمولة», never «مشمول 0» or «0 سجل»",
          "لا سجلات مشمولة" in kpi and "مشمول 0" not in kpi and "0 سجل" not in kpi, kpi[:600])
    check("confidentiality in header, footer and @page footer",
          dom.get("conf") == "سري — للاستخدام الداخلي" and dom.get("confFooter") == "سري — للاستخدام الداخلي"
          and "@bottom-center" in dom.get("pageStyle", "") and "سري — للاستخدام الداخلي" in dom.get("pageStyle", ""), dom.get("pageStyle"))
    check("font faces present in the page", dom.get("fontFaces") == 6, dom.get("fontFaces"))
    metas = " | ".join(dom.get("meta") or [])
    check("study date reads «غير مذكور في الملف» with the last-modified date as context",
          "تاريخ إعداد الدراسة غير مذكور في الملف" in metas and f"آخر تعديل للملف المصدر: {MODIFIED}" in metas, metas)
    pr = dom.get("print") or {}
    check("print: identity captions whole (no «…»)", pr.get("captions") and not any("…" in c for c in pr["captions"]), pr.get("captions"))
    check("print: no placeholder panel on paper, the sentence and trend print", pr.get("shownCharts") == ["asset-panels", "revenue-trend"], pr.get("shownCharts"))
    check("print: confidentiality mark in the printed header", pr.get("confShown") is True)
    mob = dom.get("mobile") or {}
    cells = mob.get("cells") or {}
    check("≤ 600 px: rows are cards — value, revenue, yield, status inside the screen; no header band; no side scroll",
          not mob.get("theadShown") and mob.get("wrapSW", 9e9) <= mob.get("wrapCW", 0) + 1 and mob.get("sortShown")
          and len(cells) == 4 and all(c[0] >= 0 and c[1] <= mob["vw"] and c[2] > 0 for c in cells.values()), mob)

    print("[check] check_dashboard.mjs على التشغيل الوهمي")
    rc, log = run(["node", os.path.join(S, "check_dashboard.mjs"), "--html", os.path.join(d, "dashboard.html"), "--out", d])
    check("check_dashboard passes (review, print, mobile_table, priority_order, table_detail)", rc == 0, log[-900:])
    rep = load(os.path.join(d, "08_dashboard_check.json"))
    names = {c["name"]: c for c in rep.get("checks") or []}
    check("new checks ran: mobile_table, priority_order", {"mobile_table", "priority_order"} <= set(names), sorted(names))
    pages = ((names.get("print") or {}).get("data") or {}).get("pages")
    check("print within 8 A4 pages", isinstance(pages, int) and pages <= 8, pages)
    return finish()


def finish():
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
