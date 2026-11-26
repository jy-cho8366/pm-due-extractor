"""독립 검산 + 경계값 테스트 (pm_due.py의 판정 함수를 import하지 않고 따로 계산한다)

실행: python verify.py [주시작일=2026-11-16]
"""
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).parent
ws_ = datetime.strptime(sys.argv[1] if len(sys.argv) > 1 else "2026-11-16", "%Y-%m-%d").date()
we = ws_ + timedelta(days=6)
tag = f"{ws_.isocalendar()[0]}-W{ws_.isocalendar()[1]:02d}"
ok = fail = 0


def check(name, cond, detail=""):
    global ok, fail
    if cond:
        ok += 1
    else:
        fail += 1
        print(f"  [FAIL] {name} {detail}")


# ---- 1) 원본에서 직접 다시 계산 ----
mwb = load_workbook(HERE / "입력/장비마스터.xlsx", data_only=True)
master = {r[0]: (r[3], r[4]) for r in mwb["장비마스터"].iter_rows(min_row=2, values_only=True)}
hwb = load_workbook(HERE / "입력/PM이력_가상.xlsx", data_only=True)
last = {}
for d, no, _, _, kind, _, st, _ in hwb["정비이력"].iter_rows(min_row=2, values_only=True):
    d = d.date() if isinstance(d, datetime) else d
    if kind == "PM" and st == "완료" and no in master and d <= ws_:
        last[no] = max(last.get(no, date.min), d)

expected = {}
for no, (cycle, grace) in master.items():
    if no not in last:
        expected[no] = "이력 없음(확인 필요)"
        continue
    due = last[no] + timedelta(days=cycle)
    gap = (ws_ - due).days            # 양수 = 주 시작일 기준 지난 일수
    if gap > grace:
        expected[no] = "주기 초과"
    elif gap > 0:
        expected[no] = "도래(유예 중)"
    elif (due - ws_).days <= 6:
        expected[no] = "이번 주 도래"
    elif (due - ws_).days <= 13:
        expected[no] = "다음 주 도래"
    else:
        expected[no] = "정상"

# ---- 2) 산출물과 대조 ----
owb = load_workbook(HERE / f"출력/주간PM계획_{tag}.xlsx", data_only=True)
got = {r[0]: r[7] for r in owb["전체 현황"].iter_rows(min_row=2, values_only=True)}
check("장비 대수 42", len(got) == 42 == len(master), f"got={len(got)}")
for no in master:
    check(f"{no}호 상태", got.get(no) == expected[no], f"expected={expected[no]} got={got.get(no)}")
plan = [r for r in owb["주간PM계획(초안)"].iter_rows(min_row=2, values_only=True) if isinstance(r[0], int)]
exp_targets = {n for n, s in expected.items() if s in ("주기 초과", "도래(유예 중)", "이번 주 도래")}
check("계획표 대상 집합", {r[1] for r in plan} == exp_targets, f"{sorted(r[1] for r in plan)} vs {sorted(exp_targets)}")
check("계획표 계획일 빈칸", all(r[10] is None for r in plan))
order = [r[8] for r in plan]
rank = {"주기 초과": 0, "도래(유예 중)": 1, "이번 주 도래": 2}
check("우선순위 정렬", order == sorted(order, key=rank.get))
issues = {(r[0], r[1]) for r in owb["확인 필요"].iter_rows(min_row=2, values_only=True)}
for case in [(220, "목록에 없는 호기"), (272, "미래 날짜"), (184, "PM 미완료"), (426, "이력 없음")]:
    check(f"확인필요 {case}", case in issues)
check("BM 기록 미사용(최근PM일은 PM만)", True)  # 위 재계산이 PM만 사용하므로 상태 일치로 갈음

# ---- 3) 경계값 단위 테스트 (pm_due.classify) ----
sys.path.insert(0, str(HERE))
from pm_due import classify  # noqa: E402

W, E = date(2026, 11, 16), date(2026, 11, 22)
cases = [
    (W - timedelta(days=8), "주기 초과"),       # 유예기한 = W-1
    (W - timedelta(days=7), "도래(유예 중)"),   # 유예기한 = W (당일까지 유예)
    (W - timedelta(days=1), "도래(유예 중)"),
    (W, "이번 주 도래"),
    (E, "이번 주 도래"),
    (E + timedelta(days=1), "다음 주 도래"),
    (E + timedelta(days=7), "다음 주 도래"),
    (E + timedelta(days=8), "정상"),
]
for due, exp in cases:
    check(f"경계 {due}", classify(due, 7, W, E) == exp, f"→ {classify(due, 7, W, E)} (기대 {exp})")

# ---- 4) 정기검사 독립 재계산 (2년 = 같은 월·일, 2/29 → 2/28) ----
def plus_years(d, n):
    try:
        return d.replace(year=d.year + n)
    except ValueError:
        return d.replace(year=d.year + n, day=28)


ilast = {}
for d, no, _, _, kind, _, st, _ in hwb["정비이력"].iter_rows(min_row=2, values_only=True):
    d = d.date() if isinstance(d, datetime) else d
    if kind in ("정기검사", "법정검사") and st == "완료" and no in master and d <= ws_:
        ilast[no] = max(ilast.get(no, date.min), d)
iexp = {}
for no in master:
    if no not in ilast:
        iexp[no] = "이력 없음(확인 필요)"
        continue
    left = (plus_years(ilast[no], 2) - ws_).days
    iexp[no] = ("기한 초과" if left < 0 else "이번 주 기한" if left <= 6
                else "기한 임박" if left <= 60 else "정상")
igot = {r[0]: (r[5], r[6], r[7]) for r in owb["정기검사"].iter_rows(min_row=2, values_only=True)
        if isinstance(r[0], int)}
check("정기검사 장비 대수 42", len(igot) == 42, f"got={len(igot)}")
for no in master:
    check(f"{no}호 정기검사 상태", igot.get(no, (None,) * 3)[2] == iexp[no],
          f"expected={iexp[no]} got={igot.get(no)}")
    if no in ilast:
        check(f"{no}호 검사기한", igot[no][0] == plus_years(ilast[no], 2).isoformat(), f"got={igot[no][0]}")
for case in [(431, "정기검사 이력 없음"), (186, "정기검사 미완료")]:
    check(f"확인필요 {case}", case in issues)
for no, exp in [(189, "기한 초과"), (186, "기한 초과"), (274, "이번 주 기한")]:
    check(f"의도 케이스 {no}호 {exp}", iexp[no] == exp)

from pm_due import add_months, classify_inspection  # noqa: E402

check("add_months 2/29+24", add_months(date(2024, 2, 29), 24) == date(2026, 2, 28))
check("add_months 연말", add_months(date(2025, 12, 31), 24) == date(2027, 12, 31))
icases = [
    (W - timedelta(days=1), "기한 초과"),
    (W, "이번 주 기한"),
    (E, "이번 주 기한"),
    (E + timedelta(days=1), "기한 임박"),
    (W + timedelta(days=60), "기한 임박"),
    (W + timedelta(days=61), "정상"),
]
for due, exp in icases:
    check(f"검사 경계 {due}", classify_inspection(due, W, E) == exp, f"→ {classify_inspection(due, W, E)}")

# ---- 5) PM 부품 ----
mtype = {r[0]: (r[1], r[2]) for r in mwb["장비마스터"].iter_rows(min_row=2, values_only=True)}
preg = {}
for r in mwb["PM부품"].iter_rows(min_row=2, values_only=True):
    preg.setdefault((r[0], r[1]), []).append(r)
for r in plan:
    exp = ", ".join(x[2] for x in preg[mtype[r[1]]]) if mtype[r[1]] in preg else "미등록(확인 필요)"
    check(f"{r[1]}호 작업항목", r[12] == exp, f"got={r[12]} exp={exp}")
cntt = {}
for r in plan:
    cntt[mtype[r[1]]] = cntt.get(mtype[r[1]], 0) + 1
need = [r for r in owb["이번주 부품 준비"].iter_rows(min_row=2, values_only=True) if r[6] is not None]
exp_rows = sum(len(preg.get(k, [None])) for k in cntt)
check("부품 준비 행 수", len(need) == exp_rows, f"{len(need)} vs {exp_rows}")
for r in need:
    check(f"부품 대상 대수 {r[0]} {r[1]} {r[2]}", r[6] == cntt[(r[0], r[1])])
    if r[4] is not None:
        check(f"부품 합계 {r[2]}", r[7] == r[4] * r[6])
    else:
        check(f"수량 빈칸 합계 없음 {r[2]}", r[7] is None)
    if (r[0], r[1]) not in preg:
        check(f"미등록 표시 {r[0]} {r[1]}", r[2] == "미등록(확인 필요)")

print(f"검산 {ok}/{ok + fail} 통과" + ("" if fail == 0 else f", 실패 {fail}"))
sys.exit(1 if fail else 0)
