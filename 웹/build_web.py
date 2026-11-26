"""template.html에 기본 데이터(장비마스터 + 정비이력)를 넣어 index.html을 만든다.

실행: python 웹/build_web.py [--master 입력/장비마스터.xlsx] [--history 입력/PM이력_가상.xlsx]
"""
import argparse
import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).parent
ROOT = HERE.parent


def rows(path, sheet):
    ws = load_workbook(path, data_only=True)[sheet]
    it = ws.iter_rows(values_only=True)
    head = [str(h) for h in next(it)]
    return [dict(zip(head, r)) for r in it if any(c is not None for c in r)]


def d2s(v):
    return v.strftime("%Y-%m-%d") if isinstance(v, (date, datetime)) else v


ap = argparse.ArgumentParser()
ap.add_argument("--master", default=str(ROOT / "입력/장비마스터.xlsx"))
ap.add_argument("--history", default=str(ROOT / "입력/PM이력_가상.xlsx"))
a = ap.parse_args()

master = [{"no": int(r["호기"]), "kind": r["구분"], "type": r["타입"],
           "cycle": int(r["PM주기(일)"]), "grace": int(r["유예(일)"]),
           "inspMonths": int(r.get("정기검사주기(개월)") or 24)} for r in rows(a.master, "장비마스터")]
parts = []
if "PM부품" in load_workbook(a.master, read_only=True).sheetnames:
    for r in rows(a.master, "PM부품"):
        if not r.get("작업항목"):
            continue
        q = r.get("수량(대당)")
        parts.append({"kind": str(r["구분"]).strip(), "type": str(r["타입"]).strip(),
                      "item": str(r["작업항목"]).strip(), "part": str(r.get("부품명") or "").strip(),
                      "qty": q if isinstance(q, (int, float)) else None, "unit": str(r.get("단위") or "").strip(),
                      "note": str(r.get("비고") or "").strip()})
history = [{"date": d2s(r["정비일자"]), "no": r["호기"], "pmType": r["정비유형"], "status": r["상태"]}
           for r in rows(a.history, "정비이력")]
data = {"master": master, "history": history, "parts": parts,
        "source": f"장비마스터 {Path(a.master).name} · 정비이력 {Path(a.history).name} ({len(history)}행)"}

payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
html = (HERE / "template.html").read_text(encoding="utf-8").replace("/*__DATA__*/", payload)
out = HERE / "index.html"
out.write_text(html, encoding="utf-8")
(ROOT / "index.html").write_text(html, encoding="utf-8")  # 루트 사본도 같은 내용으로
print(f"{out} + {ROOT / 'index.html'} ({len(master)}대, 이력 {len(history)}행, PM부품 {len(parts)}항목, {len(html):,} bytes)")
