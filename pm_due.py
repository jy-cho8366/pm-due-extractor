"""주간 PM 도래 장비 자동 추출기

장비마스터(호기·PM주기·유예)와 정비이력(PM 실시일)을 읽어
이번 주 PM 대상 장비를 찾고 주간 PM 계획표 초안(엑셀)과 요약(md)을 만든다.

사용법:
  python pm_due.py --master 입력/장비마스터.xlsx --history 입력/PM이력_가상.xlsx [--week-start 2026-11-16] [--out 출력]

판정 규칙 (팀_공통_전제.md 3-1):
  최근PM일 = 해당 호기의 정비유형 'PM' + 상태 '완료' 기록 중 가장 늦은 날짜 (주 시작일 이후 날짜는 오입력으로 제외)
  도래일   = 최근PM일 + PM주기
  유예기한 = 도래일 + 유예(일)
  - 유예기한 < 주 시작일              → 주기 초과
  - 도래일 < 주 시작일 ≤ 유예기한     → 도래(유예 중)
  - 주 시작일 ≤ 도래일 ≤ 주 종료일   → 이번 주 도래
  - 주 종료일 < 도래일 ≤ 종료일+7     → 다음 주 도래
  - 그 외                              → 정상
  완료 PM 기록이 없으면 '이력 없음(확인 필요)' — 추측하지 않는다.

정기검사 (2026-11-26 사용자 확인: 전 장비 2년에 1회):
  최근검사일 = 정비유형 '정기검사' 또는 '법정검사' + 상태 '완료' 기록 중 가장 늦은 날짜
  검사기한   = 최근검사일 + 정기검사주기(개월, 기본 24개월). 2/29 → 2/28
  유예 없음(0일, 가정 — 확인 필요). 임박 기준 60일(검사 신청 준비 기간 가정 — 확인 필요)
  - 검사기한 < 주 시작일                → 기한 초과
  - 주 시작일 ≤ 검사기한 ≤ 주 종료일    → 이번 주 기한
  - 주 종료일 < 검사기한 ≤ 시작일+60일  → 기한 임박
  - 그 외                                → 정상
  완료 검사 기록이 없으면 '이력 없음(확인 필요)'.
PM 부품 (2026-11-26 추가):
  장비마스터의 'PM부품' 시트(구분·타입별 작업항목·부품명·대당 수량)를 읽어
  이번 주 PM 대상 장비의 작업항목을 계획표에 붙이고, 타입별 부품 준비 수량을 합산한다.
  합계 수량 = 대당 수량 × 이번 주 대상 대수. 수량이 빈칸이면 합산하지 않는다(확인 필요).
  항목이 등록되지 않은 타입은 '미등록(확인 필요)'으로 표시한다 — 추측해서 채우지 않는다.
계획일·입항 일정·검사 신청 일정 배정은 사람이 한다(빈칸으로 둔다).
"""
import argparse
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

FONT = "맑은 고딕"
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
INPUT_FILL = PatternFill("solid", fgColor="FFF9DB")
STATUS_FILL = {
    "주기 초과": PatternFill("solid", fgColor="F8CBAD"),
    "도래(유예 중)": PatternFill("solid", fgColor="FCE4D6"),
    "이번 주 도래": PatternFill("solid", fgColor="FFF2CC"),
    "다음 주 도래": PatternFill("solid", fgColor="DDEBF7"),
}
PRIORITY = {"주기 초과": 0, "도래(유예 중)": 1, "이번 주 도래": 2, "다음 주 도래": 3, "정상": 4}

INSPECT_TYPES = {"정기검사", "법정검사"}
INSPECT_MONTHS_DEFAULT = 24   # 2년에 1회, 전 장비 동일 (2026-11-26 사용자 확인)
INSPECT_WARN_DAYS = 60        # 기한 임박 표시 기준 (가정, 확인 필요)
INSPECT_FILL = {
    "기한 초과": PatternFill("solid", fgColor="F8CBAD"),
    "이번 주 기한": PatternFill("solid", fgColor="FCE4D6"),
    "기한 임박": PatternFill("solid", fgColor="FFF2CC"),
}
INSPECT_PRIORITY = {"기한 초과": 0, "이번 주 기한": 1, "기한 임박": 2, "이력 없음(확인 필요)": 3, "정상": 4}


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        return datetime.strptime(v.strip()[:10], "%Y-%m-%d").date()
    raise ValueError(f"날짜로 읽을 수 없음: {v!r}")


def read_rows(path, sheet):
    wb = load_workbook(path, data_only=True, read_only=True)
    if sheet not in wb.sheetnames:
        sys.exit(f"[오류] {path}에 '{sheet}' 시트가 없습니다. 시트 목록: {wb.sheetnames}")
    it = wb[sheet].iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(it)]
    return [dict(zip(header, r)) for r in it if any(c is not None for c in r)]


def load_master(path):
    out = {}
    for r in read_rows(path, "장비마스터"):
        no = int(r["호기"])
        out[no] = {"호기": no, "구분": r["구분"], "타입": r["타입"],
                   "주기": int(r["PM주기(일)"]), "유예": int(r["유예(일)"]),
                   "검사주기": int(r.get("정기검사주기(개월)") or INSPECT_MONTHS_DEFAULT)}
    return out


def load_parts(path):
    """PM부품 시트 → {(구분, 타입): [항목...]}. 시트가 없으면 None."""
    wb = load_workbook(path, read_only=True)
    if "PM부품" not in wb.sheetnames:
        return None
    parts = {}
    for r in read_rows(path, "PM부품"):
        item = str(r.get("작업항목") or "").strip()
        if not item:
            continue
        qty = r.get("수량(대당)")
        if isinstance(qty, str):
            qty = float(qty) if qty.strip().replace(".", "", 1).isdigit() else None
        key = (str(r["구분"]).strip(), str(r["타입"]).strip())
        parts.setdefault(key, []).append({
            "작업항목": item, "부품명": str(r.get("부품명") or "").strip(), "수량": qty,
            "단위": str(r.get("단위") or "").strip(), "비고": str(r.get("비고") or "").strip()})
    return parts


def items_text(parts, r):
    lst = (parts or {}).get((r["구분"], r["타입"]))
    return ", ".join(dict.fromkeys(p["작업항목"] for p in lst)) if lst else "미등록(확인 필요)"


def parts_need(parts, targets):
    """이번 주 대상 장비 기준 타입별 부품 준비 목록."""
    groups = {}
    for r in targets:
        groups.setdefault((r["구분"], r["타입"]), []).append(r["호기"])
    out = []
    for (kind, typ), nos in groups.items():
        lst = (parts or {}).get((kind, typ))
        base = {"구분": kind, "타입": typ, "대상 대수": len(nos), "대상 호기": ", ".join(map(str, sorted(nos)))}
        if not lst:
            out.append({**base, "작업항목": "미등록(확인 필요)", "부품명": "", "수량": None, "단위": "",
                        "합계": None, "비고": "장비마스터 'PM부품' 시트에 이 타입의 항목을 등록해야 함"})
            continue
        for p in lst:
            total = p["수량"] * len(nos) if isinstance(p["수량"], (int, float)) else None
            note = p["비고"] or ("" if total is not None else "대당 수량 미입력(확인 필요)")
            out.append({**base, **p, "합계": total, "비고": note})
    return out


def classify(due, grace, ws, we):
    limit = due + timedelta(days=grace)
    if limit < ws:
        return "주기 초과"
    if due < ws:
        return "도래(유예 중)"
    if due <= we:
        return "이번 주 도래"
    if due <= we + timedelta(days=7):
        return "다음 주 도래"
    return "정상"


def add_months(d, n):
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    for day in (d.day, 28):  # 2/29 → 평년 2/28
        try:
            return date(y, m, day)
        except ValueError:
            continue


def classify_inspection(due, ws, we, warn=INSPECT_WARN_DAYS):
    if due < ws:
        return "기한 초과"
    if due <= we:
        return "이번 주 기한"
    if due <= ws + timedelta(days=warn):
        return "기한 임박"
    return "정상"


def analyze(master, history, ws):
    we = ws + timedelta(days=6)
    issues = []
    last_done, in_prog = {}, {}
    insp_rows = []
    for i, r in enumerate(history, start=2):
        try:
            no = int(r["호기"])
            d = to_date(r["정비일자"])
        except (TypeError, ValueError) as e:
            issues.append({"호기": r.get("호기"), "구분": "입력 오류", "내용": f"{i}행 읽기 실패: {e}"})
            continue
        kind = str(r.get("정비유형", "")).strip()
        if kind in INSPECT_TYPES:
            insp_rows.append((i, no, d, str(r.get("상태", "")).strip()))
            continue
        if kind.upper() != "PM":
            continue
        if no not in master:
            issues.append({"호기": no, "구분": "목록에 없는 호기",
                           "내용": f"{i}행 {d} PM 기록 — 장비마스터에 없는 호기(오입력 또는 신규 장비 확인 필요)"})
            continue
        if d > ws:
            issues.append({"호기": no, "구분": "미래 날짜",
                           "내용": f"{i}행 PM일 {d}이 주 시작일 {ws}보다 늦음 — 오입력으로 보고 판정에서 제외"})
            continue
        st = str(r.get("상태", "")).strip()
        if st == "완료":
            if no not in last_done or d > last_done[no]:
                last_done[no] = d
        else:
            if no not in in_prog or d > in_prog[no][0]:
                in_prog[no] = (d, st)

    result = []
    for no, m in master.items():
        row = dict(m)
        last = last_done.get(no)
        note = ""
        if no in in_prog and (last is None or in_prog[no][0] >= last):
            d, st = in_prog[no]
            note = f"{d} PM '{st}' — 완료되면 판정이 바뀜"
            issues.append({"호기": no, "구분": "PM 미완료", "내용": f"{d} PM 상태 '{st}'. 직전 완료 PM 기준으로 판정함"})
        if last is None:
            row.update(최근PM일=None, 도래일=None, 유예기한=None, 상태="이력 없음(확인 필요)", 초과일수=None, 비고=note)
            issues.append({"호기": no, "구분": "이력 없음", "내용": "완료된 PM 기록이 없음 — 대장·CMMS 확인 필요"})
        else:
            due = last + timedelta(days=m["주기"])
            st = classify(due, m["유예"], ws, we)
            over = (ws - due).days if due < ws else 0
            row.update(최근PM일=last, 도래일=due, 유예기한=due + timedelta(days=m["유예"]),
                       상태=st, 초과일수=over, 비고=note)
        result.append(row)
    result.sort(key=lambda x: (PRIORITY.get(x["상태"], 9), -(x["초과일수"] or 0),
                               x["도래일"] or date.max, x["호기"]))
    inspection = analyze_inspection(master, insp_rows, ws, we, issues)
    return result, issues, we, inspection


def analyze_inspection(master, rows, ws, we, issues):
    last_done, in_prog = {}, {}
    for i, no, d, st in rows:
        if no not in master:
            issues.append({"호기": no, "구분": "정기검사 목록에 없는 호기",
                           "내용": f"{i}행 {d} 정기검사 기록 — 장비마스터에 없는 호기(확인 필요)"})
            continue
        if d > ws:
            issues.append({"호기": no, "구분": "정기검사 미래 날짜",
                           "내용": f"{i}행 검사일 {d}이 주 시작일 {ws}보다 늦음 — 오입력으로 보고 판정에서 제외"})
            continue
        if st == "완료":
            last_done[no] = max(last_done.get(no, date.min), d)
        elif no not in in_prog or d > in_prog[no][0]:
            in_prog[no] = (d, st)

    out = []
    for no, m in master.items():
        row = {"호기": no, "구분": m["구분"], "타입": m["타입"], "검사주기": m["검사주기"], "비고": ""}
        last = last_done.get(no)
        if no in in_prog and (last is None or in_prog[no][0] >= last):
            d, st = in_prog[no]
            row["비고"] = f"{d} 정기검사 '{st}' — 완료되면 판정이 바뀜"
            issues.append({"호기": no, "구분": "정기검사 미완료",
                           "내용": f"{d} 정기검사 상태 '{st}'. 직전 완료 검사 기준으로 판정함"})
        if last is None:
            row.update(최근검사일=None, 검사기한=None, 남은일수=None, 상태="이력 없음(확인 필요)")
            issues.append({"호기": no, "구분": "정기검사 이력 없음",
                           "내용": "완료된 정기검사(법정검사) 기록이 없음 — 검사증·CMMS 확인 필요"})
        else:
            due = add_months(last, m["검사주기"])
            row.update(최근검사일=last, 검사기한=due, 남은일수=(due - ws).days,
                       상태=classify_inspection(due, ws, we))
        out.append(row)
    out.sort(key=lambda x: (INSPECT_PRIORITY.get(x["상태"], 9),
                            x["검사기한"] or date.max, x["호기"]))
    return out


def style_header(ws, widths):
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate(widths):
        ws.column_dimensions[ws.cell(row=1, column=i + 1).column_letter].width = w
    ws.freeze_panes = "A2"


def fmt(d):
    return d.isoformat() if d else ""


def write_outputs(result, issues, ws_, we, out_dir, src, inspection, parts):
    out_dir.mkdir(parents=True, exist_ok=True)
    iso = ws_.isocalendar()
    tag = f"{iso[0]}-W{iso[1]:02d}"
    wb = Workbook()

    # 1. 주간 PM 계획표 초안
    sh = wb.active
    sh.title = "주간PM계획(초안)"
    cols = ["순위", "호기", "구분", "타입", "PM주기(일)", "최근PM일", "도래일", "유예기한", "상태", "초과일수",
            "계획일(직접 입력)", "입항 일정 확인", "PM 작업항목", "비고"]
    sh.append(cols)
    targets = [r for r in result if r["상태"] in ("주기 초과", "도래(유예 중)", "이번 주 도래")]
    for i, r in enumerate(targets, 1):
        sh.append([i, r["호기"], r["구분"], r["타입"], r["주기"], fmt(r["최근PM일"]), fmt(r["도래일"]),
                   fmt(r["유예기한"]), r["상태"], r["초과일수"], None, None, items_text(parts, r), r["비고"]])
        for c in sh[sh.max_row]:
            c.font = Font(name=FONT)
        sh.cell(row=sh.max_row, column=9).fill = STATUS_FILL[r["상태"]]
        sh.cell(row=sh.max_row, column=11).fill = INPUT_FILL
        sh.cell(row=sh.max_row, column=12).fill = INPUT_FILL
    style_header(sh, [6, 7, 7, 11, 10, 12, 12, 12, 13, 9, 16, 14, 28, 40])
    sh.cell(row=sh.max_row + 2, column=1,
            value=f"대상 주: {ws_} ~ {we} ({tag}) · 노란 칸(계획일·입항 일정)은 담당자가 입력 · 원본: {src}").font = Font(name=FONT, italic=True)

    # 1-2. 이번 주 부품 준비
    need = parts_need(parts, targets)
    shp = wb.create_sheet("이번주 부품 준비")
    shp.append(["구분", "타입", "작업항목", "부품명", "대당 수량", "단위", "대상 대수", "합계 수량", "대상 호기", "비고"])
    for n in need:
        shp.append([n["구분"], n["타입"], n["작업항목"], n["부품명"], n["수량"], n["단위"],
                    n["대상 대수"], n["합계"], n["대상 호기"], n["비고"]])
        if "확인 필요" in n["작업항목"] or n["합계"] is None:
            shp.cell(row=shp.max_row, column=10).fill = STATUS_FILL["도래(유예 중)"]
    style_header(shp, [7, 11, 14, 16, 9, 6, 9, 9, 30, 44])
    shp.cell(row=shp.max_row + 2, column=1,
             value="대상 = 이번 주 PM 계획표(주기 초과·유예 중·이번 주 도래) · 재고 확인·구매 요청은 담당자가 한다"
             + ("" if parts is not None else " · 장비마스터에 'PM부품' 시트가 없음")).font = Font(name=FONT, italic=True)

    # 2. 다음 주 예고
    sh2 = wb.create_sheet("다음주 예고")
    sh2.append(["호기", "구분", "타입", "최근PM일", "도래일", "비고"])
    for r in [r for r in result if r["상태"] == "다음 주 도래"]:
        sh2.append([r["호기"], r["구분"], r["타입"], fmt(r["최근PM일"]), fmt(r["도래일"]), r["비고"]])
    style_header(sh2, [7, 7, 11, 12, 12, 40])

    # 3. 확인 필요
    sh3 = wb.create_sheet("확인 필요")
    sh3.append(["호기", "구분", "내용"])
    for it in issues:
        sh3.append([it["호기"], it["구분"], it["내용"]])
    style_header(sh3, [8, 16, 80])

    # 4. 전체 현황
    sh4 = wb.create_sheet("전체 현황")
    sh4.append(["호기", "구분", "타입", "PM주기(일)", "최근PM일", "도래일", "유예기한", "상태", "초과일수", "비고"])
    for r in sorted(result, key=lambda x: x["호기"]):
        sh4.append([r["호기"], r["구분"], r["타입"], r["주기"], fmt(r["최근PM일"]), fmt(r["도래일"]),
                    fmt(r["유예기한"]), r["상태"], r["초과일수"], r["비고"]])
        if r["상태"] in STATUS_FILL:
            sh4.cell(row=sh4.max_row, column=8).fill = STATUS_FILL[r["상태"]]
    style_header(sh4, [7, 7, 11, 10, 12, 12, 12, 16, 9, 40])

    # 5. 정기검사
    sh6 = wb.create_sheet("정기검사")
    sh6.append(["호기", "구분", "타입", "검사주기(개월)", "최근검사일", "검사기한", "남은일수", "상태",
                "검사 신청일(직접 입력)", "비고"])
    for r in inspection:
        sh6.append([r["호기"], r["구분"], r["타입"], r["검사주기"], fmt(r["최근검사일"]), fmt(r["검사기한"]),
                    r["남은일수"], r["상태"], None, r["비고"]])
        if r["상태"] in INSPECT_FILL:
            sh6.cell(row=sh6.max_row, column=8).fill = INSPECT_FILL[r["상태"]]
            sh6.cell(row=sh6.max_row, column=9).fill = INPUT_FILL
    style_header(sh6, [7, 7, 11, 12, 12, 12, 9, 18, 18, 40])
    sh6.cell(row=sh6.max_row + 2, column=1,
             value=f"기준일 {ws_} · 남은일수 = 검사기한 − 주 시작일(음수면 초과) · "
                   f"임박 기준 {INSPECT_WARN_DAYS}일·유예 0일은 가정값(확인 필요)").font = Font(name=FONT, italic=True)

    # 6. 기준
    sh5 = wb.create_sheet("판정 기준")
    for line in (__doc__ or "").strip().splitlines():
        sh5.append([line])
    sh5.column_dimensions["A"].width = 110

    xlsx = out_dir / f"주간PM계획_{tag}.xlsx"
    wb.save(xlsx)

    # 요약 md
    cnt = {}
    for r in result:
        cnt[r["상태"]] = cnt.get(r["상태"], 0) + 1
    lines = [f"# 주간 PM 도래 장비 요약 ({ws_} ~ {we}, {tag})", "",
             f"- 대상 장비 {len(result)}대 · 원본: `{src}`",
             "- 상태별: " + " / ".join(f"{k} {cnt.get(k, 0)}대" for k in
                                     ["주기 초과", "도래(유예 중)", "이번 주 도래", "다음 주 도래", "정상", "이력 없음(확인 필요)"]),
             "", "## 이번 주 PM 대상 (우선순위 순)", "",
             "| 순위 | 호기 | 구분·타입 | 최근PM일 | 도래일 | 상태 | 초과일수 | 비고 |", "|---:|---:|---|---|---|---|---:|---|"]
    for i, r in enumerate(targets, 1):
        kind = r["구분"] if r["타입"] == "-" else f"{r['구분']} {r['타입']}"
        lines.append(f"| {i} | {r['호기']} | {kind} | {fmt(r['최근PM일'])} | {fmt(r['도래일'])} | "
                     f"{r['상태']} | {r['초과일수']} | {r['비고']} |")
    lines += ["", "## 이번 주 PM 부품 준비", "",
              "| 구분·타입 | 작업항목 | 부품명 | 대당 수량 | 대상 대수 | 합계 | 비고 |",
              "|---|---|---|---:|---:|---:|---|"]
    for n in need:
        kind = n["구분"] if n["타입"] == "-" else f"{n['구분']} {n['타입']}"
        q = "" if n["수량"] is None else f"{n['수량']:g}{n['단위']}"
        t = "" if n["합계"] is None else f"{n['합계']:g}{n['단위']}"
        lines.append(f"| {kind} | {n['작업항목']} | {n['부품명']} | {q} | {n['대상 대수']} | {t} | {n['비고']} |")
    icnt = {}
    for r in inspection:
        icnt[r["상태"]] = icnt.get(r["상태"], 0) + 1
    itargets = [r for r in inspection if r["상태"] in ("기한 초과", "이번 주 기한", "기한 임박")]
    lines += ["", f"## 정기검사 (2년 주기 · 기한 초과·{INSPECT_WARN_DAYS}일 이내)", "",
              "- 상태별: " + " / ".join(f"{k} {icnt.get(k, 0)}대" for k in INSPECT_PRIORITY), ""]
    if itargets:
        lines += ["| 호기 | 구분·타입 | 최근검사일 | 검사기한 | 남은일수 | 상태 | 비고 |",
                  "|---:|---|---|---|---:|---|---|"]
        for r in itargets:
            kind = r["구분"] if r["타입"] == "-" else f"{r['구분']} {r['타입']}"
            lines.append(f"| {r['호기']} | {kind} | {fmt(r['최근검사일'])} | {fmt(r['검사기한'])} | "
                         f"{r['남은일수']} | {r['상태']} | {r['비고']} |")
    else:
        lines.append("- 기한 초과·임박 장비 없음")
    lines += ["", "## 확인 필요", ""] + [f"- **{it['호기']}** ({it['구분']}): {it['내용']}" for it in issues]
    lines += ["", "> 계획일·입항 일정 배정과 PM 연기 여부 판단은 담당자가 한다. 유예 7일은 가정값(확인 필요).",
              f"> 정기검사: 유예 0일·임박 기준 {INSPECT_WARN_DAYS}일은 가정값(확인 필요). 검사 신청·수검 일정은 담당자가 정한다."]
    md = out_dir / f"주간PM요약_{tag}.md"
    md.write_text("\n".join(lines), encoding="utf-8")
    return xlsx, md, targets, cnt, itargets


def main():
    ap = argparse.ArgumentParser(description="주간 PM 도래 장비 추출")
    ap.add_argument("--master", required=True)
    ap.add_argument("--history", required=True)
    ap.add_argument("--history-sheet", default="정비이력")
    ap.add_argument("--week-start", help="주 시작일(월요일) YYYY-MM-DD. 생략 시 이번 주 월요일")
    ap.add_argument("--out", default=str(Path(__file__).parent / "출력"))
    a = ap.parse_args()

    if a.week_start:
        try:
            ws_ = datetime.strptime(a.week_start, "%Y-%m-%d").date()
        except ValueError:
            sys.exit(f"[오류] --week-start 형식은 YYYY-MM-DD 입니다: {a.week_start}")
        if ws_.weekday() != 0:
            sys.exit(f"[오류] --week-start는 월요일이어야 합니다: {ws_} ({'월화수목금토일'[ws_.weekday()]}요일)")
    else:
        t = date.today()
        ws_ = t - timedelta(days=t.weekday())

    master = load_master(a.master)
    parts = load_parts(a.master)
    history = read_rows(a.history, a.history_sheet)
    result, issues, we, inspection = analyze(master, history, ws_)
    xlsx, md, targets, cnt, itargets = write_outputs(result, issues, ws_, we, Path(a.out),
                                                     Path(a.history).name, inspection, parts)
    print(f"대상 주 {ws_}~{we} | 이번 주 PM 대상 {len(targets)}대 | 상태별 {cnt} | "
          f"정기검사 초과·임박 {len(itargets)}대 | 확인 필요 {len(issues)}건")
    print(f"→ {xlsx}\n→ {md}")


if __name__ == "__main__":
    main()
