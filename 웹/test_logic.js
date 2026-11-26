// index.html의 판정 로직(script#logic)을 꺼내 Python 버전(pm_due.py) 산출물 엑셀의 '전체 현황'·'확인 필요'와 대조한다.
// 실행: node 웹/test_logic.js <python_result.json>
//   python_result.json = verify_web.py가 만든 {week: {statuses:{호기:상태}, issues:[[호기,구분]]}}
const fs = require('fs'), path = require('path'), vm = require('vm');
const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const logic = html.match(/<script id="logic">([\s\S]*?)<\/script>/)[1];
const data = JSON.parse(html.match(/<script id="data" type="application\/json">([\s\S]*?)<\/script>/)[1]);
const ctx = { module: { exports: {} } }; vm.createContext(ctx); vm.runInContext(logic, ctx);
const { analyze, parseDay, classify, fmtDay, addMonths, classifyInspection, itemsText, partsNeed } = ctx.module.exports;
const py = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
let ok = 0, fail = 0;
const check = (name, cond, info = '') => { if (cond) ok++; else { fail++; console.log('  [FAIL]', name, info); } };

for (const [week, exp] of Object.entries(py)) {
  const { result, issues, inspection } = analyze(data.master, data.history, parseDay(week));
  check(`${week} 대수`, result.length === Object.keys(exp.statuses).length);
  for (const r of result) check(`${week} ${r.no}호`, exp.statuses[r.no] === r.status, `py=${exp.statuses[r.no]} js=${r.status}`);
  const js = issues.map(i => `${i.no}|${i.kind}`).sort().join(',');
  const pi = exp.issues.map(i => `${i[0]}|${i[1]}`).sort().join(',');
  check(`${week} 확인 필요`, js === pi, `py=${pi} js=${js}`);
  const jsOrder = result.filter(r => ['주기 초과', '도래(유예 중)', '이번 주 도래'].includes(r.status)).map(r => r.no).join(',');
  check(`${week} 계획표 순서`, jsOrder === exp.plan.join(','), `py=${exp.plan} js=${jsOrder}`);
  // PM 작업항목·부품 준비
  const targets = result.filter(r => ['주기 초과', '도래(유예 중)', '이번 주 도래'].includes(r.status));
  for (const r of targets) check(`${week} ${r.no}호 작업항목`, itemsText(data.parts, r) === exp.items[r.no], `py=${exp.items[r.no]} js=${itemsText(data.parts, r)}`);
  const jp = partsNeed(data.parts, targets).map(n => [n.kind, n.type, n.item, n.count, n.total, n.nos].join('|')).join(';');
  const pp = exp.parts.map(n => n.join('|')).join(';');
  check(`${week} 부품 준비`, jp === pp, `\n   py=${pp}\n   js=${jp}`);
  // 정기검사
  check(`${week} 정기검사 대수`, inspection.length === Object.keys(exp.insp).length);
  for (const r of inspection) {
    const e = exp.insp[r.no] || [];
    check(`${week} ${r.no}호 정기검사`, e[0] === fmtDay(r.due) && e[1] === r.status, `py=${e} js=${fmtDay(r.due)},${r.status}`);
  }
}
check('addMonths 2/29+24', fmtDay(addMonths(parseDay('2024-02-29'), 24)) === '2026-02-28');
check('addMonths 2/29+48', fmtDay(addMonths(parseDay('2024-02-29'), 48)) === '2028-02-29');
check('addMonths 2/29+12', fmtDay(addMonths(parseDay('2024-02-29'), 12)) === '2025-02-28');
check('addMonths 연말', fmtDay(addMonths(parseDay('2025-12-31'), 24)) === '2027-12-31');
{
  const W = parseDay('2026-11-16'), E = W + 6;
  [[W - 1, '기한 초과'], [W, '이번 주 기한'], [E, '이번 주 기한'], [E + 1, '기한 임박'], [W + 60, '기한 임박'], [W + 61, '정상']]
    .forEach(([due, exp]) => check(`검사 경계 ${fmtDay(due)}`, classifyInspection(due, W, E) === exp, classifyInspection(due, W, E)));
}
// 부품 수량 합산 (가상 수량)
{
  const ps = [{ kind: 'T/C', type: '엔진발전기', item: '필터 교체', part: '필터', qty: 3, unit: '개', note: '' }];
  const n = partsNeed(ps, [{ no: 216, kind: 'T/C', type: '엔진발전기' }, { no: 277, kind: 'T/C', type: '엔진발전기' }, { no: 181, kind: 'G/C', type: '-' }]);
  check('부품 합계 3×2', n[0].total === 6 && n[0].nos === '216, 277');
  check('미등록 타입', n[1].item === '미등록(확인 필요)' && n[1].total === null);
}
// 날짜 파싱: 엑셀 일련번호·문자열·Date 동일 결과
const d = parseDay('2026-11-16');
check('엑셀 일련번호', parseDay(46342) === d, parseDay(46342) + ' vs ' + d);
check('Date 객체', parseDay(new Date(2026, 10, 16)) === d);
check('fmtDay 왕복', fmtDay(d) === '2026-11-16');
let threw = false; try { parseDay('2026/11/16'); } catch (e) { threw = true; } check('잘못된 형식 거부', threw);
// 경계값 (Python verify.py와 동일 케이스)
const W = d, E = d + 6;
[[W - 8, '주기 초과'], [W - 7, '도래(유예 중)'], [W - 1, '도래(유예 중)'], [W, '이번 주 도래'], [E, '이번 주 도래'],
 [E + 1, '다음 주 도래'], [E + 7, '다음 주 도래'], [E + 8, '정상']].forEach(([due, exp]) =>
  check(`경계 ${fmtDay(due)}`, classify(due, 7, W, E) === exp, classify(due, 7, W, E)));
console.log(`JS 로직 대조 ${ok}/${ok + fail} 통과` + (fail ? `, 실패 ${fail}` : ''));
process.exit(fail ? 1 : 0);
