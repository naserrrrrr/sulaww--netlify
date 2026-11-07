#!/usr/bin/env python3
"""Version comparison, follow-up table and delivery message on a copy of the suite's run (FICTIONAL data only).

    python3 -I tests/test_versions_message.py --run RUN --work DIR [--skip-browser]

(U1) A previous version that differs in one revenue value: compare_versions puts it in «data_correction»; with
     version_comparison and recommendation_followup in 06, the dashboard shows the five category labels and the
     follow-up table, and analysis.json carries both.
(U4) A 09_message.md with one number of 05, one wrong number and one phone: validate reports message_numbers
     (warning) for the wrong number and pii_text_patterns (blocking) for the phone.
Exit 0 = all pass, 1 = a failure.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

sys.dont_write_bytecode = True
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
S = os.path.join(SKILL, "scripts")
PY = [sys.executable, "-I"]
FAILS, PASSES = [], []
CATEGORY_LABELS = ["أداء فترة جديدة", "تغير نطاق المحفظة", "تغير التقييم", "تصحيح بيانات للفترة نفسها",
                   "تغير تعريف أو تصنيف"]

DOM_PROBE = r"""
import { createRequire } from 'module';
import { execSync } from 'child_process';
import path from 'path';
let pw = null;
for (const root of [process.cwd() + '/x.js', path.join(execSync('npm root -g', {encoding: 'utf8'}).trim(), 'x.js')]) {
  try { pw = createRequire(root)('playwright'); break; } catch (e) { /* next */ }
}
if (!pw) { console.log(JSON.stringify({error: 'playwright'})); process.exit(0); }
const b = await pw.chromium.launch();
const p = await b.newPage();
await p.goto('file://' + process.argv[2]);
await p.waitForTimeout(500);
const r = await p.evaluate(() => {
  document.querySelectorAll('details').forEach((d) => { d.open = true; });
  const sec = document.querySelector('[data-section="versions"]') || document.getElementById('sec-versions') || document.body;
  return { text: sec.innerText, followup: !!document.querySelector('[data-role="followup-table"] table, [data-role="followup-table"]'),
           cats: [...document.querySelectorAll('[data-role="version-category"]')].map((e) => e.textContent) };
});
console.log(JSON.stringify(r));
await b.close();
"""


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {str(info)[:400]}"))


def run(cmd, **kw):
    p = subprocess.run(cmd, capture_output=True, text=True, **kw)
    return p.returncode, p.stdout + p.stderr


def load(p):
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def dump(p, o):
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(o, fh, ensure_ascii=False, indent=1)


def copy_run(src, dst):
    shutil.rmtree(dst, ignore_errors=True)
    os.makedirs(dst)
    for f in ("01_workbook_map.json", "02_mapping.json", "03_records.json", "04_metrics_config.json", "05_metrics.json",
              "06_narrative.json"):
        if os.path.isfile(os.path.join(src, f)):
            shutil.copy(os.path.join(src, f), os.path.join(dst, f))


def versions(run_dir, work, skip_browser):
    print("[U1] مقارنة الإصدارات ومتابعة التوصيات في الداشبورد")
    prev, cur = os.path.join(work, "prev"), os.path.join(work, "curr")
    copy_run(run_dir, prev)
    copy_run(run_dir, cur)
    recs = load(os.path.join(prev, "03_records.json"))
    meta = recs.get("field_meta") or {}
    rev = next(f for f, m in meta.items() if (m or {}).get("group") == "revenue" and f in
               (load(os.path.join(prev, "05_metrics.json"))["metrics"]["property_revenue"].get("resolved_fields") or {}).values())
    target = next(r for r in recs["records"] if r.get("scope") == "in_scope"
                  and (r["fields"].get(rev) or {}).get("s") == "value")
    target["fields"][rev]["v"] = target["fields"][rev]["v"] - 10000   # the earlier version recorded 10,000 less
    dump(os.path.join(prev, "03_records.json"), recs)
    rc, log = run(PY + [os.path.join(S, "compute_metrics.py"), "--records", os.path.join(prev, "03_records.json"),
                        "--config", os.path.join(prev, "04_metrics_config.json"), "--out", os.path.join(prev, "05_metrics.json")])
    check("previous version recomputed", rc == 0, log[-300:])
    out = os.path.join(cur, "version_comparison.json")
    rc, log = run(PY + [os.path.join(S, "compare_versions.py"), "--prev", prev, "--curr", cur, "--out", out])
    check("compare_versions exits 0", rc == 0, log[-300:])
    if rc != 0:
        return
    vc = load(out)
    cats = vc.get("categories") or vc.get("changes") or {}
    corr = cats.get("data_correction") or []
    check("the changed value is a data correction of the same period",
          any(c.get("record_id_curr") == target["id"] and c.get("field") == rev for c in corr), corr[:3])
    nar = load(os.path.join(cur, "06_narrative.json"))
    nar["version_comparison"] = vc
    rid = (nar.get("recommendations") or [{}])[0].get("id")
    nar["recommendation_followup"] = [{"rec_id": rid, "status": "قيد التنفيذ", "evidence": None,
                                       "note": "بدأ حصر الحالة ولم يكتمل"}]
    dump(os.path.join(cur, "06_narrative.json"), nar)
    rc, log = run(PY + [os.path.join(S, "validate_analysis.py"), "--run-dir", cur])
    check("validate with version_comparison: no blocking error", rc in (0, 1), log[-400:])
    rc, log = run(PY + [os.path.join(S, "build_dashboard.py"), "--run-dir", cur])
    check("build_dashboard with version_comparison", rc == 0, log[-400:])
    if rc != 0:
        return
    an = load(os.path.join(cur, "analysis.json"))
    nar_an = an.get("narrative") or an
    check("analysis.json carries version_comparison and recommendation_followup",
          isinstance(nar_an.get("version_comparison"), dict) and len(nar_an.get("recommendation_followup") or []) == 1,
          list(nar_an.keys())[:20])
    if skip_browser:
        print("  SKIP  DOM: versions section (APA_SKIP_BROWSER=1)")
        return
    rc, log = run(["node", os.path.join(S, "check_dashboard.mjs"), "--html", os.path.join(cur, "dashboard.html"),
                   "--out", cur])
    check("check_dashboard passes with a non-empty versions section", rc == 0, log[-600:])
    probe = os.path.join(work, "versions_probe.mjs")
    with open(probe, "w", encoding="utf-8") as fh:
        fh.write(DOM_PROBE)
    rc, log = run(["node", probe, os.path.join(cur, "dashboard.html")], cwd=work)
    try:
        dom = json.loads(log.strip().splitlines()[-1])
    except (ValueError, IndexError):
        dom = {"error": log[-300:]}
    if dom.get("error"):
        check("DOM probe ran", False, dom)
        return
    text = dom.get("text") or ""
    check("versions section shows the five category labels", all(lab in text for lab in CATEGORY_LABELS),
          [lab for lab in CATEGORY_LABELS if lab not in text])
    check("versions section shows the follow-up table with its status", dom.get("followup") and "قيد التنفيذ" in text,
          dom.get("followup"))


def message(run_dir, work):
    print("[U4] رسالة التسليم 09_message.md")
    d = os.path.join(work, "msg")
    copy_run(run_dir, d)
    m05 = load(os.path.join(d, "05_metrics.json"))["metrics"]
    good = next(r["value"] for r in m05.values() if r.get("unit") == "SAR" and isinstance(r.get("value"), (int, float))
                and r["value"] > 10000)
    msg = ("السلام عليكم،\n"
           f"بلغت الإيرادات المسجلة {good:,.0f} ريال وفق الملف.\n"
           "وتبلغ قيمة الأصول غير المقيمة 7,654,321 ريال.\n"
           "للتواصل: 0551234567\n")
    with open(os.path.join(d, "09_message.md"), "w", encoding="utf-8") as fh:
        fh.write(msg)
    rc, log = run(PY + [os.path.join(S, "validate_analysis.py"), "--run-dir", d])
    check("validate with a phone in 09_message.md exits 2", rc == 2, log[-300:])
    v = load(os.path.join(d, "07_validation.json"))
    failed = [c for c in v["checks"] if not c.get("passed")]
    mn = [c for c in failed if c["id"] == "message_numbers"]
    check("message_numbers warning names the wrong number only",
          mn and all(c["level"] == "warning" for c in mn) and any("7,654,321" in c["detail"] for c in mn)
          and not any(f"{good:,.0f}" in c["detail"] for c in mn), [c["detail"] for c in mn])
    pii = [c for c in failed if c["id"] == "pii_text_patterns" and "09_message" in str(c.get("where"))]
    check("pii_text_patterns blocks the phone in the message", pii and all(c["level"] == "blocking" for c in pii),
          [(c["id"], c["level"], c.get("where")) for c in failed][:8])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--work", required=True)
    ap.add_argument("--skip-browser", action="store_true")
    a = ap.parse_args()
    os.makedirs(a.work, exist_ok=True)
    versions(a.run, a.work, a.skip_browser or os.environ.get("APA_SKIP_BROWSER") == "1")
    message(a.run, a.work)
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
