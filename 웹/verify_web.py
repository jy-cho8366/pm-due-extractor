"""웹 버전 검증: Python 추출기를 여러 주로 실행 → 결과를 JSON으로 → node로 JS 로직과 대조.
테스트 산출물은 웹/_test 에 만들고 끝나면 지운다.

실행: python 웹/verify_web.py
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).parent
ROOT = HERE.parent
TMP = HERE / "_test"
NODE = r"C:\Program Files\nodejs\node.exe"
WEEKS = ["2026-11-02", "2026-11-09", "2026-11-16", "2026-11-23", "2026-11-30", "2026-12-14"]

TMP.mkdir(exist_ok=True)
out = {}
try:
    for w in WEEKS:
        subprocess.run([sys.executable, str(ROOT / "pm_due.py"), "--master", str(ROOT / "입력/장비마스터.xlsx"),
                        "--history", str(ROOT / "입력/PM이력_가상.xlsx"), "--week-start", w, "--out", str(TMP)],
                       check=True, capture_output=True)
        f = next(TMP.glob("주간PM계획_*.xlsx"))
        wb = load_workbook(f, data_only=True)
        out[w] = {
            "statuses": {str(r[0]): r[7] for r in wb["전체 현황"].iter_rows(min_row=2, values_only=True)},
            "issues": [[r[0], r[1]] for r in wb["확인 필요"].iter_rows(min_row=2, values_only=True)],
            "plan": [r[1] for r in wb["주간PM계획(초안)"].iter_rows(min_row=2, values_only=True) if isinstance(r[0], int)],
            "items": {str(r[1]): r[12] for r in wb["주간PM계획(초안)"].iter_rows(min_row=2, values_only=True)
                      if isinstance(r[0], int)},
            "insp": {str(r[0]): [r[5] or "", r[7]] for r in wb["정기검사"].iter_rows(min_row=2, values_only=True)
                     if isinstance(r[0], int)},
            "parts": [[r[0], r[1], r[2], r[6], r[7], r[8]] for r in wb["이번주 부품 준비"].iter_rows(min_row=2, values_only=True)
                      if r[6] is not None],
        }
        wb.close()
        for p in TMP.iterdir():
            p.unlink()
    j = TMP / "py.json"
    j.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    r = subprocess.run([NODE, str(HERE / "test_logic.js"), str(j)], capture_output=True, text=True, encoding="utf-8")
    print(r.stdout.strip(), r.stderr.strip())
    code = r.returncode
finally:
    shutil.rmtree(TMP, ignore_errors=True)
sys.exit(code)
