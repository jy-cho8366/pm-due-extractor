"""입력 파일 2개를 만든다.

1) 장비마스터.xlsx  : 실제 보유 장비 42대(사용자 제공, 2026-11-19)와 PM 주기
2) PM이력_가상.xlsx : 테스트용 **가상** 정비 이력 (실제 데이터가 아님)

실행: python make_inputs.py
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

HERE = Path(__file__).parent
DATA = HERE / "입력"
DATA.mkdir(exist_ok=True)

# ---- 실제 장비 목록 (팀_공통_전제.md 1-1, 3-1) ----
GC = list(range(181, 191))
TC_ENGINE = [216, 217, 218, 219, 221, 228, 229, 277, 278]
TC_ELEC = [215, 224, 225, 226, 227, 271, 272, 273, 274, 275, 276, 279,
           422, 424, 425, 426, 428, 430, 431, 432, 433, 434, 435]
TYPES = [  # (구분, 타입, 호기목록, 주기일)
    ("G/C", "-", GC, 45),
    ("T/C", "엔진발전기", TC_ENGINE, 30),
    ("T/C", "전기", TC_ELEC, 60),
]
GRACE = 7  # 유예일 (가정, 확인 필요)
INSPECT_MONTHS = 24  # 정기검사 2년에 1회, 전 장비 동일 (2026-11-26 사용자 확인)

# PM 작업항목·부품 (2026-11-26 사용자 예시: 엔진장비 = 필터 교체, 윤활유 교체)
# 수량·필터 종류는 아직 모름 → 빈칸. G/C·전기 T/C 항목은 받지 않았으므로 등록하지 않는다(확인 필요).
PM_PARTS = [  # (구분, 타입, 작업항목, 부품명, 수량(대당), 단위, 비고)
    ("T/C", "엔진발전기", "필터 교체", "필터", None, "개", "필터 종류(엔진오일·연료·에어 등)·수량 확인 필요"),
    ("T/C", "엔진발전기", "윤활유 교체", "윤활유", None, "L", "규격·교체량 확인 필요"),
]

HEAD_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")


def write_sheet(ws, header, rows, widths):
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A2"


def make_master():
    wb = Workbook()
    ws = wb.active
    ws.title = "장비마스터"
    rows = []
    for kind, typ, nums, cycle in TYPES:
        for n in nums:
            rows.append([n, kind, typ, cycle, GRACE, INSPECT_MONTHS])
    write_sheet(ws, ["호기", "구분", "타입", "PM주기(일)", "유예(일)", "정기검사주기(개월)"], rows,
                [8, 8, 12, 12, 10, 18])
    ws2 = wb.create_sheet("안내")
    ws2["A1"] = "장비 목록과 PM 주기는 2026-11-19 사용자 제공 값입니다. 유예 7일은 가정값(확인 필요)입니다."
    ws2["A2"] = "정기검사 주기 24개월(2년 1회, 전 장비 동일)은 2026-11-26 사용자 확인 값입니다."
    ws2["A3"] = ("PM부품 시트: 구분·타입별로 한 줄에 작업항목 하나. 타입은 장비마스터와 같은 값(G/C는 '-')을 쓴다. "
                 "수량(대당)을 채우면 이번 주 준비 수량이 합산된다. G/C·전기 T/C 항목은 아직 미등록(확인 필요).")
    wp = wb.create_sheet("PM부품", 1)
    write_sheet(wp, ["구분", "타입", "작업항목", "부품명", "수량(대당)", "단위", "비고"], [list(r) for r in PM_PARTS],
                [8, 12, 14, 14, 11, 6, 44])
    wb.save(DATA / "장비마스터.xlsx")
    return rows


def make_history(master, base: date):
    """기준일 근처에 도래·초과·정상 장비가 골고루 나오도록 가상 이력을 만든다."""
    rng = random.Random(20261119)
    rows = []
    workers = ["정비1반", "정비2반", "협력업체A"]
    pm_tasks = {"G/C": "정기점검(와이어로프·브레이크·스프레더)",
                "엔진발전기": "정기점검(엔진오일·연료필터·발전기)",
                "전기": "정기점검(급전장치·모터·브레이크)"}
    skip_history = {426}          # 이력 없음 케이스
    in_progress = {184}           # 최근 PM이 진행중 케이스
    for no, kind, typ, cycle, *_ in master:
        if no in skip_history:
            continue
        key = kind if kind == "G/C" else typ
        # 마지막 PM이 기준일로부터 '주기 ± 랜덤' 전이 되도록 → 일부는 도래, 일부는 초과
        offset = rng.randint(-12, cycle - 3)
        last = base - timedelta(days=cycle - offset)
        # 그 이전 이력 2회
        for k in (2, 1):
            d = last - timedelta(days=cycle * k + rng.randint(-3, 3))
            rows.append([d, no, kind, typ, "PM", pm_tasks[key], "완료", rng.choice(workers)])
        status = "진행중" if no in in_progress else "완료"
        rows.append([last, no, kind, typ, "PM", pm_tasks[key], status, rng.choice(workers)])
        if no in in_progress:  # 진행중 이전의 완료 PM 하나 더 (그 기준으로 판정)
            pass
        # 고장수리(BM) 기록 섞기 — PM 판정에 쓰면 안 됨
        if rng.random() < 0.35:
            d = base - timedelta(days=rng.randint(1, 40))
            rows.append([d, no, kind, typ, "BM", "고장수리(가상)", "완료", rng.choice(workers)])
    # 입력 오류 케이스: 목록에 없는 호기, 미래 날짜
    rows.append([base - timedelta(days=10), 220, "T/C", "전기", "PM", "정기점검", "완료", "정비1반"])
    rows.append([base + timedelta(days=30), 272, "T/C", "전기", "PM", "정기점검(날짜 오입력)", "완료", "정비2반"])
    rows += make_inspections(master, base)
    rows.sort(key=lambda r: (r[0], r[1]))
    wb = Workbook()
    ws = wb.active
    ws.title = "정비이력"
    write_sheet(ws, ["정비일자", "호기", "구분", "타입", "정비유형", "작업내용", "상태", "담당"],
                rows, [12, 8, 8, 12, 10, 36, 8, 12])
    for r in ws.iter_rows(min_row=2, max_col=1):
        r[0].number_format = "yyyy-mm-dd"
    ws2 = wb.create_sheet("안내")
    ws2["A1"] = "※ 테스트용 가상 이력입니다. 실제 정비 이력이 아닙니다."
    ws2["A2"] = "의도한 테스트 케이스: 426호 이력 없음 / 184호 최근 PM 진행중 / 220호 목록에 없는 호기 / 272호 미래 날짜 오입력 / BM 기록 혼입"
    ws2["A3"] = ("정기검사 테스트 케이스: 189호 기한 초과 / 186호 기한 초과 + 재검사 진행중 / 274호 이번 주 기한 / "
                 "431호 검사 이력 없음 / 일부 기록은 정비유형 '법정검사'로 표기")
    wb.save(DATA / "PM이력_가상.xlsx")
    return len(rows)


def years_ago(d, n):
    try:
        return d.replace(year=d.year - n)
    except ValueError:  # 2/29
        return d.replace(year=d.year - n, day=28)


def make_inspections(master, base: date):
    """정기검사(2년 1회) 가상 기록. PM 난수와 분리해 기존 PM 이력이 바뀌지 않게 한다."""
    rng = random.Random(20261126)
    fixed = {  # 호기: 최근 완료 검사일 (의도한 케이스)
        189: years_ago(base, 2) - timedelta(days=10),   # 기한 초과
        186: years_ago(base, 2) - timedelta(days=5),    # 기한 초과 + 재검사 진행중
        274: years_ago(base, 2) + timedelta(days=3),    # 이번 주 기한
    }
    rows = []
    for no, kind, typ, _, _, _ in master:
        if no == 431:  # 검사 이력 없음
            continue
        last = fixed.get(no) or base - timedelta(days=rng.randint(70, 700))
        label = "법정검사" if rng.random() < 0.3 else "정기검사"
        for d in (years_ago(last, 2), last):
            rows.append([d, no, kind, typ, label, "정기검사 수검(가상)", "완료", "검사기관"])
    rows.append([base - timedelta(days=3), 186, "G/C", "-", "정기검사", "정기검사 수검(가상)", "진행중", "검사기관"])
    return rows


if __name__ == "__main__":
    m = make_master()
    n = make_history(m, date(2026, 11, 16))
    print(f"장비마스터 {len(m)}대, 가상 이력 {n}행 → {DATA}")
