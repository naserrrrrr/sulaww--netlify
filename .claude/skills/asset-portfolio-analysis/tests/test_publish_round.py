#!/usr/bin/env python3
"""Shareable-page and audit round, on a copy of the suite's run (FICTIONAL data only).

    python3 -I tests/test_publish_round.py --run RUN --xlsx SYNTHETIC.xlsx --work DIR

(P1) validate pii_values reads every row of a workbook saved without a <dimension> record (read-only openpyxl then
     reports max_row = None): the check runs (not skipped), passes on clean outputs and blocks a personal value copied
     into the narrative.
(P2) build_dashboard --publish: records outside the scope are withheld (one summary row per sheet, no names), deed
     numbers are masked to their last three digits everywhere, and no local path (/home/, /tmp/) reaches the page in any
     build.
(P3) export_workbook: every workbook anomaly row of «مشكلات البيانات» states its impact and treatment; «الحسابات»
     freezes columns only; a reason never reads as a raw status token; counts, concentrations and coverage are
     recomputed with live formulas; plan items carry a follow-up date; ratio values keep the dashboard's precision.
(P4) validate warns on an «assumption» while no scenario is prepared (assumption_kinds).
(P5) inspect_workbook never copies a link into the column samples of 01.
(P6) the page's cell-reference test is linear: a long list of references that does not match («…، … (+5)») is
     rejected at once (one regular expression over the whole list backtracked exponentially and froze the page).
Exit 0 = all pass, 1 = a failure.
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

sys.dont_write_bytecode = True
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
S = os.path.join(SKILL, "scripts")
PY = [sys.executable, "-I"]
FAILS, PASSES = [], []


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {str(info)[:400]}"))


def run(cmd, **kw):
    p = subprocess.run(cmd, capture_output=True, text=True, **kw)
    return p.returncode, p.stdout + p.stderr


def copy_run(src, dst):
    if os.path.isdir(dst):
        shutil.rmtree(dst)
    os.makedirs(dst)
    for f in os.listdir(src):
        p = os.path.join(src, f)
        if os.path.isfile(p) and (f.endswith(".json") or f.endswith(".md")):
            shutil.copy(p, os.path.join(dst, f))


def strip_dimensions(src, dst):
    """The same workbook without any <dimension> element in its sheet XML (as some producers save it)."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                data = re.sub(rb"<dimension\b[^>]*/>", b"", data)
            zout.writestr(item, data)


def checks_of(run_dir):
    v = json.load(open(os.path.join(run_dir, "07_validation.json"), encoding="utf-8"))
    return {c["id"]: c for c in v["checks"]}


def p1_pii_values(run_dir, xlsx, work):
    import openpyxl
    d = os.path.join(work, "p1")
    copy_run(run_dir, d)
    nodim = os.path.join(work, "p1_nodim.xlsx")
    strip_dimensions(xlsx, nodim)
    with zipfile.ZipFile(nodim) as z:
        n_dim = sum(z.read(n).count(b"<dimension") for n in z.namelist() if n.startswith("xl/worksheets/"))
    wb = openpyxl.load_workbook(nodim, read_only=True)
    ws = wb[wb.sheetnames[0]]
    check("P1 fixture: no <dimension> record, read-only max_row is None", n_dim == 0 and ws.max_row is None, (n_dim, ws.max_row))
    wb.close()
    # the scan must still see the personal columns' values: a clean run passes (not skipped) ...
    run(PY + [os.path.join(S, "validate_analysis.py"), "--run-dir", d, "--xlsx", nodim])
    c = checks_of(d).get("pii_values") or {}
    check("P1 pii_values runs on a workbook without <dimension> (not skipped) and passes", c.get("passed") and not c.get("skipped"), c)
    # ... and a value of the personal-data column copied into the narrative blocks
    wb = openpyxl.load_workbook(xlsx)
    w = json.load(open(os.path.join(d, "01_workbook_map.json"), encoding="utf-8"))
    pc = (w.get("pii_columns") or [{}])[0]
    ws = wb[pc["sheet"]]
    col = pc["col"]
    val = next((ws[f"{col}{r}"].value for r in range(3, ws.max_row + 1) if ws[f"{col}{r}"].value), None)
    n = json.load(open(os.path.join(d, "06_narrative.json"), encoding="utf-8"))
    n.setdefault("methodology_notes", []).append(f"ملاحظة تجريبية تذكر {val}")
    json.dump(n, open(os.path.join(d, "06_narrative.json"), "w", encoding="utf-8"), ensure_ascii=False)
    run(PY + [os.path.join(S, "validate_analysis.py"), "--run-dir", d, "--xlsx", nodim])
    c = checks_of(d).get("pii_values") or {}
    check("P1 a personal value in the narrative is caught on that workbook", c.get("passed") is False and not c.get("skipped"),
          {k: c.get(k) for k in ("passed", "skipped", "detail")})


def p2_publish(run_dir, work):
    d = os.path.join(work, "p2")
    copy_run(run_dir, d)
    rc, out = run(PY + [os.path.join(S, "build_dashboard.py"), "--run-dir", d])
    html = open(os.path.join(d, "dashboard.html"), encoding="utf-8").read()
    check("P2 default build: no local path in the page (05 inputs reduced to file names)", rc == 0 and not re.search(r"/home/|/tmp/", html),
          re.findall(r".{20}(?:/home/|/tmp/).{20}", html)[:2])
    rec = json.load(open(os.path.join(d, "03_records.json"), encoding="utf-8"))
    out_scope = [r for r in rec["records"] if r.get("scope") != "in_scope"]
    deeds = sorted({str((r.get("fields") or {}).get("deed_no", {}).get("raw") or "") for r in rec["records"]} - {""})
    deed_digits = [x for x in deeds if re.fullmatch(r"\d{5,}", x)]
    rc, out = run(PY + [os.path.join(S, "build_dashboard.py"), "--run-dir", d, "--publish"])
    an = json.load(open(os.path.join(d, "analysis.json"), encoding="utf-8"))
    html = open(os.path.join(d, "dashboard.html"), encoding="utf-8").read()
    data = html.split("const DATA = ", 1)[1] if "const DATA = " in html else ""
    check("P2 --publish: only in-scope records in the page", rc == 0 and an["meta"].get("publish") is True
          and all(r.get("scope") == "in_scope" for r in an["records"]), out[-300:])
    wh = an.get("withheld") or {}
    check("P2 --publish: withheld records summarised per sheet with their count", bool(out_scope) and sum(x["count"] for x in wh.get("records", [])) == len(out_scope)
          and all(x.get("sheet") and x.get("reason") for x in wh.get("records", [])), wh)
    names = [r["name"] for r in out_scope if r.get("name") and len(r["name"]) > 6]
    leaked = [n for n in names if n in data and not any(n in (r.get("name") or "") for r in an["records"])]
    check("P2 --publish: no name of a withheld record in the page data", not leaked, leaked[:3])
    shown = [x for x in deed_digits if re.search(r"(?<!\d)" + re.escape(x) + r"(?!\d)", data)]
    check("P2 --publish: deed numbers masked to their last 3 digits", bool(deed_digits) and not shown
          and any("•••" + x[-3:] in data for x in deed_digits), (len(deed_digits), shown[:2]))


def p3_export(run_dir, work):
    import openpyxl
    d = os.path.join(work, "p3")
    copy_run(run_dir, d)
    out = os.path.join(d, "supporting_workbook.xlsx")
    rc, log = run(PY + [os.path.join(S, "export_workbook.py"), "--run-dir", d])
    check("P3 export runs", rc == 0 and os.path.isfile(out), log[-400:])
    if not os.path.isfile(out):
        return
    wb = openpyxl.load_workbook(out)
    ws = wb["مشكلات البيانات"]
    hdr = next(r for r in ws.iter_rows(values_only=True) if r and r[0] == "المعرّف")
    hi, ti, si = hdr.index("الأثر"), hdr.index("المعالجة"), hdr.index("المصدر")
    wb_rows = [r for r in ws.iter_rows(values_only=True) if r and r[si] == "فحص بنية الملف"]
    empty = [r[0] for r in wb_rows if not r[hi] or not r[ti]]
    check("P3 every workbook anomaly states its impact and treatment", wb_rows and not empty, empty[:5])
    check("P3 «الحسابات» freezes columns only (the header row of one table never sits over another)",
          wb["الحسابات"].freeze_panes == "C1", wb["الحسابات"].freeze_panes)
    check("P3 «الفجوات والتوصيات» has no row freeze", wb["الفجوات والتوصيات"].freeze_panes in (None, "A1"), wb["الفجوات والتوصيات"].freeze_panes)
    raw = [c.coordinate for ws in wb.worksheets for row in ws.iter_rows() for c in row
           if isinstance(c.value, str) and re.search(r"\((?:blank|absent|a_unknown|b_unknown|num_unknown|den_unknown)\)", c.value)]
    check("P3 no raw status token in a reason", not raw, raw[:5])
    calc = wb["الحسابات"]
    labels = [str(c.value) for row in calc.iter_rows() for c in row if isinstance(c.value, str)]
    has_conc = any(x.startswith("حصة أكبر") for x in labels)
    has_cnt = any(x.startswith("العدد (COUNTA") for x in labels)
    has_cov = any(x.startswith("التغطية (أدنى نسبة") for x in labels)
    m05 = json.load(open(os.path.join(d, "05_metrics.json"), encoding="utf-8"))["metrics"]
    kinds = {r.get("kind") for r in m05.values() if r.get("status") in ("complete", "partial")}
    check("P3 counts, concentrations and coverage recomputed with live formulas",
          (has_cnt or "count" not in kinds) and (has_conc or "concentration" not in kinds) and (has_cov or "coverage" not in kinds),
          (has_cnt, has_conc, has_cov, sorted(k for k in kinds if k)))
    pct_cells = [c for row in wb["ملخص"].iter_rows() for c in row if isinstance(c.value, (int, float)) and '"%"' in str(c.number_format)]
    bad = [(c.coordinate, c.value, c.number_format) for c in pct_cells if abs(c.value) < 10 and c.number_format != '0.00"%"']
    check("P3 ratio values keep the dashboard precision (two decimals below 10)", not bad, bad[:3])
    plan = wb["الخطة التشغيلية"]
    ph = next(r for r in plan.iter_rows(values_only=True) if r and r[0] == "المرحلة")
    li = ph.index("تاريخ المتابعة")
    m2 = json.load(open(os.path.join(d, "02_mapping.json"), encoding="utf-8"))
    rows = [r for r in plan.iter_rows(values_only=True) if r and r[1] and str(r[1]).startswith("R")]
    if m2.get("analysis_date"):
        check("P3 plan items carry a proposed follow-up date", rows and all(r[li] is not None for r in rows), [r[li] for r in rows][:3])
    for r in rows:
        if "نافذة المرحلة" not in str(r[ph.index("المدة")]) and not r[ph.index("المدة")]:
            check("P3 plan duration filled", False, r)
            break


def p4_assumption_kind(run_dir, work):
    d = os.path.join(work, "p4")
    copy_run(run_dir, d)
    n = json.load(open(os.path.join(d, "06_narrative.json"), encoding="utf-8"))
    n.setdefault("scenarios", {})["status"] = "not_prepared"
    n.setdefault("assumptions", []).append({"id": "A99", "text": "افتراض تجريبي", "kind": "assumption", "basis": "اختيار منهجي تجريبي"})
    json.dump(n, open(os.path.join(d, "06_narrative.json"), "w", encoding="utf-8"), ensure_ascii=False)
    run(PY + [os.path.join(S, "validate_analysis.py"), "--run-dir", d])
    c = checks_of(d).get("assumption_kinds") or {}
    check("P4 an «assumption» without prepared scenarios is reported", c.get("passed") is False and c.get("level") == "warning", c)


def p5_inspect_links(work):
    import openpyxl
    d = os.path.join(work, "p5")
    os.makedirs(d, exist_ok=True)
    x = os.path.join(d, "links.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "سجل"
    ws.append(["الاسم", "الإيراد", "الموقع على الخريطة"])
    for i in range(1, 6):
        ws.append([f"عقار تجريبي {i}", 1000 * i, f"https://maps.example.invalid/pin{i}"])
    wb.save(x)
    rc, out = run(PY + [os.path.join(S, "inspect_workbook.py"), x, "--out", d])
    blob = open(os.path.join(d, "01_workbook_map.json"), encoding="utf-8").read() if os.path.isfile(os.path.join(d, "01_workbook_map.json")) else ""
    check("P5 no link copied into the workbook map samples", rc == 0 and "maps.example.invalid" not in blob, out[-300:])


REF_PROBE = r"""
const fs = require('fs');
const src = fs.readFileSync(process.argv[2], 'utf8');
const pick = (name) => { const m = src.match(new RegExp('\\n  var ' + name + ' = [\\s\\S]*?;\\n')); return m ? m[0] : ''; };
let code = ['REF1', 'SHEET1', 'REF_PART_RE', 'REF_FIRST_RE'].map(pick).join('');
const m = src.match(/\n  var REF_ONLY_RE = \{[\s\S]*?\} \};\n/);
code += m ? m[0] : '';
const REF_ONLY_RE = new Function(code + '; return REF_ONLY_RE;')();
const sheet = 'ورقة تجريبية ';
const long = Array.from({ length: 40 }, (_, i) => sheet + '!N' + (i + 3)).join(', ') + ', … (+5)';
const ok = Array.from({ length: 40 }, (_, i) => sheet + '!N' + (i + 3)).join(', ');
const t = Date.now();
const r1 = REF_ONLY_RE.test(long.replace(/\s+!/g, '!')), r2 = REF_ONLY_RE.test(ok.replace(/\s+!/g, '!'));
console.log(JSON.stringify({ ms: Date.now() - t, nonmatch: r1, match: r2, found: !!m }));
"""


def p6_ref_regex(work):
    if not shutil.which("node"):
        print("  SKIP  P6 (node غير مثبت)")
        return
    probe = os.path.join(work, "ref_probe.cjs")
    open(probe, "w", encoding="utf-8").write(REF_PROBE)
    try:
        p = subprocess.run(["node", probe, os.path.join(SKILL, "assets", "dashboard_template.html")], capture_output=True,
                           text=True, timeout=20)
        res = json.loads(p.stdout.strip().splitlines()[-1]) if p.stdout.strip() else {"error": p.stderr[-300:]}
    except subprocess.TimeoutExpired:
        res = {"timeout": True}
    check("P6 a long non-matching reference list is rejected in linear time", res.get("found") and res.get("nonmatch") is False
          and res.get("match") is True and res.get("ms", 9999) < 1000, res)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--work", required=True)
    a = ap.parse_args()
    if os.path.isdir(a.work):
        shutil.rmtree(a.work)
    os.makedirs(a.work)
    p1_pii_values(a.run, a.xlsx, a.work)
    p2_publish(a.run, a.work)
    p3_export(a.run, a.work)
    p4_assumption_kind(a.run, a.work)
    p5_inspect_links(a.work)
    p6_ref_regex(a.work)
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
