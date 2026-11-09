#!/usr/bin/env python3
"""Supporting workbook: the formulas recalculate to the values the export stores, on FICTIONAL data only.

    python3 -I tests/test_workbook_export.py      # exit 0 = pass (or LibreOffice absent: SKIP), 1 = a failure

The export writes live formulas (SUM per field, the check «مطابق/فرق», the ratio, the change in riyals, the live
differences of «المطابقة مع الإجماليات», and «ملخص» links to «الحسابات») and stores the value each one evaluates to,
so a preview that does not recalculate still shows the numbers. Here:
  1. export() stores a value for every formula and no formula refers to its own cell;
  2. the same workbook built WITHOUT stored values is recalculated by LibreOffice (headless conversion), and every
     formula gives the stored value, with no error cell (Err:…, #REF!, #VALUE!, #NAME?, #DIV/0!) — a review once found
     a SUM whose range was its own cell (a circular reference) when a metric had no included record.
The fictional run is the one of `export_workbook.py --self-test` (a sum with no included record, a ratio, an unverified
growth, a reviewed metric, totals matched in scope, by section, by original values, and unexplained).
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.environ.get("APA_SKILL_DIR") or os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SKILL, "scripts"))
sys.dont_write_bytecode = True

import openpyxl  # noqa: E402

import export_workbook as ex  # noqa: E402

FAILS, PASSES = [], []
ERR_PREFIXES = ("Err:", "#REF", "#VALUE", "#NAME", "#DIV", "#N/A", "#NUM", "#NULL")


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {info}"))


def same(a, b):
    if ex.is_num(a) or ex.is_num(b):
        return ex.is_num(a) and ex.is_num(b) and abs(a - b) <= 1e-6 * max(1.0, abs(b))
    return (a or "") == (b or "")


def main():
    base = os.environ.get("APA_BUILD_DIR")
    work = os.path.join(base, "workbook_export") if base else tempfile.mkdtemp(prefix="apa_wb_")
    shutil.rmtree(work, ignore_errors=True)
    run = os.path.join(work, "run")
    os.makedirs(run)
    try:
        ex._st_run(run)
        out = os.path.join(run, "supporting_workbook.xlsx")
        print("[1] القيم المخزنة للصيغ والمراجع الدائرية")
        rc = ex.export(run, out)
        check("export exits 0", rc == 0 and os.path.isfile(out), rc)
        if rc != 0:
            return 1
        wf = openpyxl.load_workbook(out)
        wv = openpyxl.load_workbook(out, data_only=True)
        formulas = [(ws.title, c.coordinate) for ws in wf.worksheets for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith("=")]
        check("the workbook has live formulas (sums, checks, links, differences)", len(formulas) >= 20, len(formulas))
        missing = [f"{s}!{c}" for s, c in formulas if wv[s][c].value is None]
        check("every formula stores its value", not missing, missing[:8])
        check("no formula refers to its own cell", ex.formula_self_refs(out) == [], ex.formula_self_refs(out))

        print("[2] إعادة الحساب في LibreOffice")
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            print("  SKIP  LibreOffice غير مثبت؛ لم تُختبر إعادة الحساب")
            return 1 if FAILS else 0
        raw = os.path.join(work, "nocache.xlsx")
        ex.PRIVACY.reset()
        exp = ex.Export(run)
        exp.build(raw)   # formulas only, no stored values: LibreOffice must calculate every one
        lo_dir = os.path.join(work, "lo")
        os.makedirs(lo_dir)
        env = dict(os.environ, HOME=os.path.join(work, "lohome"))
        try:
            subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                            "--outdir", lo_dir, raw], check=True, capture_output=True, env=env, timeout=240)
        except Exception as exc:  # pragma: no cover
            check("LibreOffice converts the workbook", False, exc)
            return 1
        lo = openpyxl.load_workbook(os.path.join(lo_dir, "nocache.xlsx"), data_only=True)
        bad = [(s, c, v, lo[s][c].value) for (s, c), v in exp.formula_cache.items() if not same(lo[s][c].value, v)]
        check("LibreOffice recalculation equals the stored value of every formula", not bad, bad[:6])
        errs = [f"{ws.title}!{c.coordinate}={c.value}" for ws in lo.worksheets for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith(ERR_PREFIXES)]
        check("no error cell after recalculation (no Err:522 circular reference)", not errs, errs[:6])
        stored = {(s, c): wv[s][c].value for s, c in formulas}
        check("the stored values are the ones LibreOffice computes", all(same(stored[k], exp.formula_cache.get(k))
                                                                        for k in stored if k in exp.formula_cache), "")
    finally:
        if not base:
            shutil.rmtree(work, ignore_errors=True)
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
