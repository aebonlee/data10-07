// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 실제 공정불량 Raw data 는 아직 받지 않았습니다. 여기서는 가상 표와 예시 데이터(협력사 A~F)로 검사합니다.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const E = require('../js/ai-endpoint.js');
const XLSX = require('../vendor/xlsx.full.min.js');
const XO = require('../js/xlsx-out.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
let idn = 0;
const newId = () => 't' + (++idn);
const TYPES = L.defaultTypes();

console.log('날짜·주차');
test('여러 날짜 표기 · 엑셀 일련번호 · 없는 날짜는 빈 값', () => {
  assert.equal(L.parseDate('2026-10-05'), '2026-10-05');
  assert.equal(L.parseDate('2026.10.5'), '2026-10-05');
  assert.equal(L.parseDate('2026/10/05 13:20'), '2026-10-05');
  assert.equal(L.parseDate('2026년 10월 5일'), '2026-10-05');
  assert.equal(L.parseDate('26.10.05'), '2026-10-05');
  assert.equal(L.parseDate('20261005'), '2026-10-05');
  assert.equal(L.parseDate(46300), '2026-10-05');
  assert.equal(L.parseDate('2026-02-30'), '');
  assert.equal(L.parseDate('확인 중'), '');
});
test('주차 — 목요일이 든 달 기준 (시안의 「2026년 10월 5주차」)', () => {
  assert.equal(L.weekInfo('2026-10-26').label, '2026년 10월 5주차');
  assert.equal(L.weekInfo('2026-11-01').label, '2026년 10월 5주차'); // 일요일
  assert.equal(L.weekInfo('2026-09-28').label, '2026년 10월 1주차'); // 10/1(목)이 든 주
  assert.equal(L.weekInfo('2026-11-02').label, '2026년 11월 1주차');
  assert.equal(L.weekInfo('2027-01-01').label, '2026년 12월 5주차'); // 12/31(목)이 든 주
  assert.equal(L.mondayOf('2026-10-25'), '2026-10-19');
});

console.log('엑셀 열 찾기');
test('제목 줄 아래 머리행 · 다른 이름(업체명·품 번·불량 내용)도 연결', () => {
  const d = L.detectColumns(Sample.aoa());
  assert.equal(d.headerRow, 2);
  assert.deepEqual(d.missing, []);
  assert.equal(d.map.supplier, 2); assert.equal(d.map.partNo, 3); assert.equal(d.map.problem, 5); assert.equal(d.map.qty, 6); assert.equal(d.map.action, 7);
});
test('불량유형 칸은 「불량유형」으로(내용과 따로) · 조치 칸 · 품번 칸은 없어도 읽음', () => {
  const aoa = [['발생일자', '협력사', '불량유형', '불량내용', '조치사항'], ['2026-10-05', '협력사 A', '용접 불량', '비드 끝 크랙', '재작업'], ['2026-10-06', '협력사 A', '미조임', '', '']];
  const p = L.parseSheet(aoa);
  assert.deepEqual(p.det.missing, []);
  assert.equal(p.det.map.type, 2); assert.equal(p.det.map.problem, 3);
  assert.equal(p.records[0].typeIn, '용접 불량');
  assert.equal(p.records[0].action, '재작업');
  assert.equal(p.records[1].problem, '미조임'); // 내용이 비면 유형 칸을 내용으로
  const e = L.enrich(p.records, TYPES);
  assert.equal(e[0].type, '용접불량'); assert.ok(e[0].column); // 띄어쓰기가 달라도 사전 이름에 맞춤
  assert.equal(e[1].type, '미조임');
});
test('필수 열이 없으면 행을 읽지 않고 · 사용자가 열을 고르면 다시 읽음', () => {
  const aoa = [['발생일자', '품번', '불량내용', '거래선'], ['2026-10-05', 'X-1', '균열', '협력사 A']];
  const d = L.detectColumns(aoa);
  assert.deepEqual(d.missing, ['supplier']);
  assert.equal(L.parseSheet(aoa).records.length, 0);
  const p = L.parseSheet(aoa, { overrides: { supplier: 3 } });
  assert.equal(p.records.length, 1);
  assert.equal(p.records[0].supplier, '협력사 A');
});
test('빈 줄·합계 줄 건너뜀 · 날짜 못 읽은 행 수 셈 · 칸 안 줄바꿈 정리', () => {
  const aoa = [['발생일자', '품명', '품번', '문제점', '협력사'], ['2026-10-05', '커버', 'X-1', '도장\n흐름', '협력사 B'], [], ['합계', null, null, null, null], ['미상', '커버', 'X-2', '찍힘', '협력사 B']];
  const p = L.parseSheet(aoa);
  assert.equal(p.records.length, 2);
  assert.equal(p.records[0].problem, '도장 흐름');
  assert.equal(p.badDate, 1);
  assert.equal(p.skipped, 1);
});
test('여러 시트 중 Raw data 시트를 고름', () => {
  const pick = L.pickSheet([{ name: '요약', aoa: [['협력사별 합계'], ['A', 3]] }, { name: 'RAW', aoa: Sample.aoa() }]);
  assert.equal(pick.name, 'RAW');
});
test('같은 행은 한 번만 (발생일자+협력사+품번+문제점, 띄어쓰기 무시)', () => {
  const a = { date: '2026-10-05', supplier: '협력사 A', partNo: 'x-1', problem: '용접 균열' };
  const m1 = L.mergeRecords([], [a], 0);
  const m2 = L.mergeRecords(m1.records, [{ date: '2026-10-05', supplier: '협력사A', partNo: 'X-1', problem: '용접균열' }, { ...a, date: '2026-10-06' }], m1.seq);
  assert.equal(m2.added, 1); assert.equal(m2.dup, 1); assert.equal(m2.records.length, 2);
  assert.equal(m2.records[1].id, 'd2');
});

console.log('불량유형 분류');
test('기본 사전 = 수강생이 정한 불량유형 8종(순서 그대로)', () => {
  assert.deepEqual(L.typeNames(TYPES), ['위치/치수', '간섭/정렬', '이종/누락', '용접불량', '누유/누수', '미조임', '성능이상', '기타']);
  assert.ok(TYPES.every(t => t.actions.length >= 3 && t.keywords.length > 0));
});
test('키워드로 분류 · 맞은 키워드를 근거로 남김', () => {
  const c = L.classify('브래킷 용접부 갈라짐', TYPES);
  assert.equal(c.type, '용접불량');
  assert.deepEqual(c.hits, ['용접', '갈라짐']);
  assert.equal(L.classify('이종품 혼입 (타 품번 섞임)', TYPES).type, '이종/누락');
  assert.equal(L.classify('와셔 누락', TYPES).type, '이종/누락');
  assert.equal(L.classify('홀 위치 치수 공차 초과', TYPES).type, '위치/치수');
  assert.equal(L.classify('조립 시 상대 부품과 간섭', TYPES).type, '간섭/정렬');
  assert.equal(L.classify('피팅부 오일 샘', TYPES).type, '누유/누수');
  assert.equal(L.classify('호스 클램프 미조임', TYPES).type, '미조임');
  assert.equal(L.classify('밸브 작동 불량', TYPES).type, '성능이상');
  assert.equal(L.classify('도장 흐름 발생', TYPES).type, '기타');
});
test('여러 유형이 걸리면 긴 키워드 → 같으면 사전에서 위에 있는 유형', () => {
  assert.equal(L.classify('이종 볼트 혼용', TYPES).type, '이종/누락'); // 혼용·볼트 모두 2글자 → 위의 이종/누락
  assert.equal(L.classify('볼트 체결 토크 부족', TYPES).type, '미조임');
  assert.equal(L.classify('도막 두께 부족', TYPES).type, '기타'); // 「도막 두께」(4글자)가 위치/치수의 「두께」를 이김
  assert.equal(L.classify('Crack 발생', TYPES).type, '용접불량'); // 대소문자 무시
  const custom = [{ name: '외관', keywords: ['찍힘'] }, { name: '포장', keywords: ['포장 찍힘'] }];
  assert.equal(L.classify('포장 찍힘 발견', custom).type, '포장'); // 긴 키워드(띄어쓰기 빼고 4글자)가 이김
});
test('맞는 키워드가 없으면 미분류 · 직접 고른 유형이 먼저 · 유형 이름은 구분 기호 무시', () => {
  assert.equal(L.classify('검사성적서 미첨부', TYPES).type, L.UNCLASSIFIED);
  const r = { problem: '검사성적서 미첨부', typeManual: '기타' };
  assert.deepEqual(L.typeOf(r, TYPES), { type: '기타', hits: [], manual: true });
  assert.equal(L.typeOf({ problem: '용접 기공', typeManual: '사전에 없는 유형' }, TYPES).type, '용접불량');
  assert.equal(L.findTypeName('위치·치수', TYPES), '위치/치수');
  assert.equal(L.findTypeName('누유 / 누수', TYPES), '누유/누수');
  assert.equal(L.findTypeName('용접 불량', TYPES), '용접불량');
  assert.equal(L.findTypeName('도장', TYPES), '');
  assert.equal(L.typeOf({ problem: '도장 흐름', typeIn: '모르는 코드' }, TYPES).type, '기타'); // 맞는 이름이 없으면 키워드로
});
test('설정 화면 글 ↔ 대책 목록 · 키워드 쉼표 나누기', () => {
  const a = L.parseActionLines('원인 조사 | 협력사 | 0 | 2\n\n검사 강화 | | x | 0\n유효성 확인', '부품품질팀');
  assert.deepEqual(a, [{ text: '원인 조사', dept: '협력사', start: 0, weeks: 2 }, { text: '검사 강화', dept: '부품품질팀', start: 0, weeks: 2 }, { text: '유효성 확인', dept: '부품품질팀', start: 0, weeks: 2 }]);
  assert.deepEqual(L.parseActionLines(L.formatActionLines(a)), a);
  assert.deepEqual(L.parseKeywords('균열, 크랙,\n균열 , crack'), ['균열', '크랙', 'crack']);
});

console.log('분석·개선 필요 항목');
const REC = L.mergeRecords([], Sample.records(), 0).records;
const ALL = L.enrich(REC, TYPES);
test('예시 데이터 — 협력사 6곳, 미분류 3건, 파레토 누적 100%', () => {
  const an = L.analyze(ALL, {}, {});
  assert.equal(an.suppliers.length, 6);
  assert.equal(an.unclassified, 3);
  assert.equal(an.byType[an.byType.length - 1].type, L.UNCLASSIFIED); // 미분류는 맨 뒤
  assert.equal(an.byType[an.byType.length - 1].cum, 100);
  assert.equal(an.byType.reduce((s, t) => s + t.count, 0), an.total);
  assert.equal(an.months[0].month, '2026-01');
  assert.equal(an.months[an.months.length - 1].month, '2026-10');
});
test('월별 추이는 빈 달도 0 으로 채움 · 기간 거르기', () => {
  const rows = L.enrich([{ date: '2026-01-10', supplier: 'S', partNo: 'P', problem: '찍힘' }, { date: '2026-04-02', supplier: 'S', partNo: 'P', problem: '찍힘' }], TYPES);
  const an = L.analyze(rows, {}, {});
  assert.deepEqual(an.months.map(m => m.count), [1, 0, 0, 1]);
  assert.equal(L.analyze(rows, { from: '2026-02-01' }, {}).total, 1);
});
test('심어 둔 신호 — 협력사 C 용접불량 증가 추세(1위) · 반복 품번 TP-3301', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const i = an.issues[0];
  assert.equal(i.type, '용접불량');
  assert.ok(i.flags.trend, '증가 추세');
  assert.ok(i.recent >= 7 && i.recent >= i.prev * 1.5, '최근 ' + i.recent + ' / 직전 ' + i.prev);
  assert.ok(i.repeatParts.some(p => p.partNo === 'TP-3301'));
});
test('심어 둔 신호 — 협력사 A 이종/누락·TP-1102, 협력사 D 위치/치수·TP-4405', () => {
  const a = L.analyze(ALL, { supplier: '협력사 A' }, {}).issues.find(x => x.type === '이종/누락');
  assert.ok(a && a.flags.repeat && a.repeatParts[0].partNo === 'TP-1102');
  const d = L.analyze(ALL, { supplier: '협력사 D' }, {}).issues.find(x => x.type === '위치/치수');
  assert.ok(d && d.flags.repeat && d.repeatParts.some(p => p.partNo === 'TP-4405'));
});
test('협력사 F 는 건수가 적어 규칙에 걸리지 않음', () => {
  assert.equal(L.analyze(ALL, { supplier: '협력사 F' }, {}).issues.length, 0);
});
test('규칙 경계 — 반복 기준 딱 맞으면 걸리고 하나 모자라면 안 걸림 · 미분류는 항목에서 뺌', () => {
  const mk = (n, problem) => Array.from({ length: n }, (_, i) => ({ date: '2026-0' + (1 + i % 3) + '-1' + i, supplier: 'S', partNo: 'P' + i, problem }));
  const rules = { repeatMin: 3, partRepeatMin: 2, trendMonths: 3, trendRate: 50, trendMin: 9 };
  assert.equal(L.analyze(L.enrich(mk(3, '찍힘'), TYPES), {}, rules).issues.length, 1);
  assert.equal(L.analyze(L.enrich(mk(2, '찍힘'), TYPES), {}, rules).issues.length, 0);
  assert.equal(L.analyze(L.enrich(mk(5, '성적서 누락 안내'), TYPES.filter(t => t.name !== '이종/누락')), {}, rules).issues.length, 0);
});
test('증가 추세 — 직전 3개월보다 50% 이상 늘고 최근 3건 이상', () => {
  const r = (d) => ({ date: d, supplier: 'S', partNo: 'P' + d, problem: '도장 박리' });
  const base = ['2026-04-03', '2026-05-03'].map(r);
  const rules = { repeatMin: 99, partRepeatMin: 2, trendMonths: 3, trendRate: 50, trendMin: 3 };
  const up = L.analyze(L.enrich(base.concat(['2026-07-01', '2026-08-01', '2026-09-01'].map(r)), TYPES), {}, rules);
  assert.equal(up.issues.length, 1); // 3건 ≥ 2 × 1.5
  const flat = L.analyze(L.enrich(base.concat(['2026-07-01', '2026-08-01'].map(r)), TYPES), {}, rules);
  assert.equal(flat.issues.length, 0); // 2건 < 3건(최소)
  assert.equal(L.cleanRules({ repeatMin: -1, trendMonths: 99 }).repeatMin, 5);
});

console.log('Master Plan 초안·간트');
test('초안 — 개선 필요 유형마다 사전의 기본 대책 · 유형마다 1주씩 늦춰 시작', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 C', baseDate: '2026-10-26', newId });
  const types = [...new Set(rows.map(r => r.type))];
  assert.deepEqual(types, an.issues.slice(0, 3).map(i => i.type));
  const crack = rows.filter(r => r.type === '용접불량');
  assert.equal(crack.length, TYPES.find(t => t.name === '용접불량').actions.length);
  const k = types.indexOf('용접불량');
  assert.equal(crack[0].planStart, L.addDays('2026-10-26', 7 * k));
  assert.equal(crack[0].planEnd, L.addDays(crack[0].planStart, 13)); // 2주
  assert.ok(rows.every(r => r.status === 'open' && r.source === 'template' && r.supplier === '협력사 C'));
  assert.ok(crack[0].note.includes('증가 추세'));
});
test('규칙에 걸린 항목이 없으면 건수 상위 유형으로 초안', () => {
  const an = L.analyze(ALL, { supplier: '협력사 F' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 F', baseDate: '2026-10-26', newId });
  assert.ok(rows.length > 0);
  assert.ok(rows[0].note.includes('규칙 기준 미만'));
  assert.ok(rows.every(r => r.type !== L.UNCLASSIFIED));
});
test('추진 집계 — 완료/미완료/보류 · 계획일 지난 미완료 · 추진 기간 일수', () => {
  const rows = [
    { status: 'done', planStart: '2026-10-26', planEnd: '2026-11-08' },
    { status: 'open', planStart: '2026-11-02', planEnd: '2026-11-15' },
    { status: 'hold', planStart: '2026-11-09', planEnd: '2026-12-20' },
    { status: 'open', planStart: '2026-11-30', planEnd: '2026-12-13' }
  ];
  const s = L.planStats(rows, '2026-11-20');
  assert.deepEqual([s.total, s.done, s.open, s.hold, s.overdue, s.days, s.rate], [4, 1, 2, 1, 1, 56, 25]);
});
test('간트 주 칸 — 월요일부터 · 위 줄은 달 · 칸 걸침', () => {
  const rows = [{ planStart: '2026-10-28', planEnd: '2026-11-10' }];
  const cols = L.ganttColumns(rows, 'week', { minCols: 1 });
  assert.equal(cols[0].start, '2026-10-26');
  assert.equal(cols[0].full, '2026년 10월 5주차');
  assert.equal(cols.length, 3);
  assert.deepEqual(L.groupSpans(cols), [{ group: '2026년 10월', n: 1 }, { group: '2026년 11월', n: 2 }]);
  assert.deepEqual(L.span('2026-10-28', '2026-11-10', cols), [0, 2]);
  assert.deepEqual(L.span('2026-11-03', '', cols), [1, 1]); // 끝이 비면 시작 칸 하나
  assert.equal(L.span('', '2026-11-03', cols), null);
  assert.equal(L.ganttColumns(rows, 'week', {}).length, 10); // 최소 10칸
});
test('간트 월 칸 — 위 줄은 연도 · 해를 넘김', () => {
  const cols = L.ganttColumns([{ planStart: '2026-11-15', planEnd: '2027-02-03' }], 'month', { minCols: 1 });
  assert.deepEqual(cols.map(c => c.label), ['11월', '12월', '1월', '2월']);
  assert.deepEqual(L.groupSpans(cols).map(g => g.group + g.n), ['2026년2', '2027년2']);
  assert.equal(cols[1].end, '2026-12-31');
});
test('입력 점검 — 날짜 순서 · 완료에는 실적 완료일', () => {
  assert.deepEqual(L.checkPlanRow({ action: 'a', planStart: '2026-11-02', planEnd: '2026-11-01', status: 'open' }), ['계획 완료일이 시작일보다 빠릅니다.']);
  assert.equal(L.checkPlanRow({ action: 'a', status: 'done' }).length, 1);
  assert.equal(L.checkPlanRow({ action: 'a', status: 'done', actStart: '2026-11-02', actEnd: '2026-11-05' }).length, 0);
});
test('계획 확정 — 버전 쌓기 · 실적 입력은 수정으로 보지 않고 계획일을 바꾸면 수정 있음', () => {
  const plan = { rows: [L.makeRow({ supplier: 'S', type: '기타', action: '포장 개선', planStart: '2026-11-02', planEnd: '2026-11-15' }, newId)] };
  assert.equal(L.planStateText(plan), '미확정 초안');
  const s1 = L.confirmPlan(plan, '2026-10-20T10:30:00');
  assert.equal(s1.version, 1);
  plan.rows[0].actStart = '2026-11-03'; plan.rows[0].status = 'done';
  assert.equal(L.planChangedSince(plan), false);
  assert.equal(s1.rows[0].actStart, ''); // 확정본은 복사본
  plan.rows[0].planEnd = '2026-11-22';
  assert.equal(L.planChangedSince(plan), true);
  assert.equal(L.planStateText(plan), '확정 v1 (2026-10-20 10:30) 이후 계획 수정 있음');
  assert.equal(L.confirmPlan(plan, '2026-10-21T09:00:00').version, 2);
});

console.log('AI 프롬프트·답 읽기');
const AL = L.updateAliases({}, REC);
test('별칭 — 같은 이름은 늘 같은 별칭 · 새 이름에만 다음 번호', () => {
  assert.equal(AL.suppliers['협력사 A'], '업체1');
  const again = L.updateAliases(AL, REC.concat([{ supplier: '협력사 Z', partNo: 'NEW-1' }]));
  assert.equal(again.suppliers['협력사 A'], '업체1');
  assert.equal(again.suppliers['협력사 Z'], '업체7');
  assert.equal(again.parts['NEW-1'], '품목' + (Object.keys(AL.parts).length + 1));
});
test('프롬프트 — 협력사·품번이 남지 않음(문제점 글 속 품번 포함) · 사실 숫자 포함', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 C', baseDate: '2026-10-26', newId });
  const info = L.buildAiPromptInfo({ supplier: '협력사 C', an, rows, types: TYPES, depts: L.DEFAULT_DEPTS, baseDate: '2026-10-26', aliases: AL });
  assert.deepEqual(info.leaks, []);
  assert.ok(!/협력사 [A-F]/.test(info.prompt));
  assert.ok(!/TP-\d{4}/.test(info.prompt));
  assert.ok(info.prompt.includes('협력사: 업체3'));
  assert.ok(info.prompt.includes('공정불량 ' + an.total + '건'));
  assert.ok(info.prompt.includes('| 불량유형 | 개선대책 | 담당부서 |'));
  const open = L.buildAiPromptInfo({ supplier: '협력사 C', an, rows, types: TYPES, baseDate: '2026-10-26', aliases: null });
  assert.ok(open.prompt.includes('협력사 C') && open.prompt.includes('TP-3301'));
  assert.ok(L.maskText('TP-3301 재발, TP-33011 은 다른 품번', AL).startsWith(AL.parts['TP-3301'] + ' 재발, TP-33011'));
});
const P3301 = AL.parts['TP-3301'];
test('AI 답 읽기 — 요약·표·확인 필요 · 별칭 되돌림 · 유형 이름 맞춤 · 주 → 날짜', () => {
  const ans = [
    '**요약:** 업체3 은 용접불량이 최근 늘었고 ' + P3301 + ' 에서 반복돼.',
    '',
    '| 불량유형 | 개선대책 | 담당부서 | 시작(기준일부터 몇 주 뒤) | 기간(주) |',
    '|:---|---|---|---|---|',
    '| 용접 불량 | ' + P3301 + ' 용접 예열 온도 기록 시작 | 협력사 | 0 | 2주 |',
    '| 위치·치수 | PT 검사<br>샘플링 2배 | 부품품질팀 | 1주 | 4 |',
    '| 새 유형 | 기타 대책 | | | |',
    '| 잘린 줄 |',
    '확인 필요:',
    '- 예열 설비가 있는지',
    '- 품목 3개 공정 순서'
  ].join('\n');
  const p = L.parseAiPlan(ans, { types: TYPES, aliases: AL, supplier: '협력사 C', baseDate: '2026-10-26', newId });
  assert.equal(p.summary, '협력사 C 은 용접불량이 최근 늘었고 TP-3301 에서 반복돼.');
  assert.equal(p.rows.length, 3);
  assert.equal(p.skipped, 1);
  assert.equal(p.rows[0].type, '용접불량');
  assert.equal(p.rows[1].type, '위치/치수'); // 「위치·치수」 → 사전 이름
  assert.equal(p.rows[0].action, 'TP-3301 용접 예열 온도 기록 시작');
  assert.equal(p.rows[0].planEnd, '2026-11-08');
  assert.equal(p.rows[1].action, 'PT 검사 샘플링 2배');
  assert.equal(p.rows[1].planStart, '2026-11-02');
  assert.equal(p.rows[2].type, '새 유형');
  assert.equal(p.rows[2].dept, '부품품질팀');
  assert.ok(p.rows.every(r => r.source === 'ai' && r.supplier === '협력사 C'));
  assert.deepEqual(p.checks, ['예열 설비가 있는지', '품목 3개 공정 순서']); // 「품목 3개」는 별칭이 아님
});
test('엑셀에서 복사한 탭 구분 줄도 읽음', () => {
  const p = L.parseAiPlan('기타\t도막 측정 추가\t부품품질팀\t2\t3', { types: TYPES, aliases: {}, supplier: 'S', baseDate: '2026-10-26', newId });
  assert.equal(p.rows.length, 1);
  assert.equal(p.rows[0].planStart, '2026-11-09');
  assert.equal(p.rows[0].planEnd, '2026-11-29');
});
test('AI 연결 — 키 없이 보내면 인증 머리글 없음 · https 페이지에서 http 주소 막음', () => {
  const r = E.buildChatRequest({ baseUrl: 'https://llm.example.com/v1/chat/completions/', model: 'm' }, E.promptMessages('x'));
  assert.equal(r.url, 'https://llm.example.com/v1/chat/completions');
  assert.equal(r.init.headers.Authorization, undefined);
  assert.ok(E.validateConfig({ baseUrl: 'http://10.0.0.5:8000/v1', model: 'm' }, 'https:').errors.length > 0);
  assert.equal(E.toStored({ baseUrl: 'https://a/v1', model: 'm', apiKey: 'sk-x', remember: false }).apiKey, '');
});

console.log('보고서·저장');
const SHEETS = ['요약', '불량 목록', '유형별·협력사별 집계', 'Master Plan', '간트 일정', '확정 이력'];
function reportFixture() {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 C', baseDate: '2026-10-26', newId });
  rows[0].actStart = '2026-10-26'; rows[0].actEnd = '2026-10-30'; rows[0].status = 'done';
  rows[1].status = 'hold';
  const plan = { baseDate: '2026-10-26', from: '', to: '', rows, confirmed: [] };
  L.confirmPlan(plan, '2026-10-20T11:00:00');
  return { an, rows, book: L.reportBook({ suppliers: ['협력사 C'], records: ALL, plans: { '협력사 C': plan }, rules: {}, today: '2026-10-20', types: TYPES }) };
}
test('보고서 6시트 — 화면과 같은 표: 불량 목록 · 유형별·협력사별 집계 · Master Plan · 간트(○ 계획 · ● 실적)', () => {
  const { an, rows, book } = reportFixture();
  const sh = book.sheets;
  assert.deepEqual(Object.keys(sh), SHEETS);
  assert.deepEqual(sh['불량 목록'][0], ['No', '발생일자', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치', '분류 근거']);
  assert.equal(sh['불량 목록'].length, an.total + 1);
  const ag = sh['유형별·협력사별 집계'];
  const cross = ag.findIndex(r => r[0] === '협력사' && r[1] === '위치/치수');
  assert.deepEqual(ag[cross].slice(1, 9), L.typeNames(TYPES)); // 교차표 머리 = 8종 순서
  assert.equal(ag[cross + 1][ag[cross].length - 1], an.total);
  assert.ok(ag.some(r => r[0] === '합계' && r[1] === an.total));
  const mp = sh['Master Plan'];
  assert.equal(mp.length, rows.length + 1);
  assert.equal(mp[1][10], '완료'); assert.equal(mp[2][10], '보류');
  const g = sh['간트 일정'];
  assert.equal(g[0][5], '2026년 10월'); assert.equal(g[1][5], '5주');
  assert.equal(g[2][5], '●');             // 첫 활동 실적
  const k = book.weeks.findIndex(w => w.start <= rows[1].planStart && rows[1].planStart <= w.end);
  assert.ok(k > 0);
  assert.equal(g[3][5 + k - 1], '');      // 둘째 활동은 계획 시작 주부터
  assert.equal(g[3][5 + k], '○');
  assert.deepEqual(sh['확정 이력'][1], ['협력사 C', 'v1', '2026-10-20 11:00', rows.length]);
  // 서식 표시: 상태 칸·간트 칸 색
  assert.ok(book.meta['간트 일정'].cells.some(c => c.r === 2 && c.c === 5 && c.style === 'done'));
  assert.ok(book.meta['간트 일정'].cells.some(c => c.r === 3 && c.c === 5 + k && c.style === 'plan'));
  assert.ok(book.meta['Master Plan'].cells.some(c => c.r === 2 && c.c === 10 && c.style === 'hold'));
});
test('Master Plan 이 없는 협력사도 불량 목록·집계는 나옴', () => {
  const sh = L.reportSheets({ suppliers: ['협력사 A', '협력사 B'], records: ALL, plans: {}, rules: {}, today: '2026-10-20', types: TYPES });
  assert.equal(sh['불량 목록'].length, ALL.filter(r => r.supplier === '협력사 A' || r.supplier === '협력사 B').length + 1);
  assert.equal(sh['Master Plan'][1][1], '(아직 Master Plan 이 없습니다)');
  assert.ok(sh['요약'].some(r => r[0] === '협력사 B' && r[10] === 'Master Plan 없음'));
});
test('Excel 쓰기 — 머리글 굵게·상태 색·날짜는 엑셀 날짜(yyyy-mm-dd)·틀 고정 · 다시 읽으면 값 그대로', () => {
  const { book } = reportFixture();
  const u8 = XO.build(XLSX, book.sheets, book.meta);
  const wb = XLSX.read(u8, { type: 'array', cellStyles: true });
  assert.deepEqual(wb.SheetNames, SHEETS);
  const list = wb.Sheets['불량 목록'];
  assert.equal(list.B2.t, 'n'); assert.equal(list.B2.z, 'yyyy-mm-dd');
  assert.equal(list.B2.w, book.sheets['불량 목록'][1][1]);
  const z = XLSX.CFB.read(u8, { type: 'array' });
  const xml = p => Buffer.from(z.FileIndex[z.FullPaths.indexOf('Root Entry/' + p)].content).toString('utf8');
  const styles = xml('xl/styles.xml');
  assert.ok(/<font><b\/>/.test(styles), '굵은 글꼴');
  assert.ok(styles.includes('FF1F9D55') && styles.includes('FF8E44AD'), '완료 녹색·보류 보라');
  const s2 = xml('xl/worksheets/sheet2.xml');
  const headId = (/<c r="A1" s="(\d+)"/.exec(s2) || [])[1];
  assert.ok(headId && +headId >= 2, '불량 목록 머리글에 서식');
  assert.ok(s2.includes('state="frozen"'), '틀 고정');
  assert.ok(s2.includes('<autoFilter'), '거르기');
  const g = xml('xl/worksheets/sheet5.xml');
  assert.ok(/<c r="F3" s="\d+"/.test(g), '간트 실적 칸 색');
  const back = XLSX.utils.sheet_to_json(wb.Sheets['Master Plan'], { header: 1, raw: false });
  assert.equal(back[1][10], '완료');
  assert.equal(back[1][6], book.sheets['Master Plan'][1][6]); // 계획 시작일 글자 그대로
});
test('저장 형식 — 깨진 값은 버리고 기본값 · 사용자 사전·계획 유지', () => {
  assert.equal(L.restoreDb(null).types.length, 8);
  const db = L.restoreDb({
    records: [{ date: '2026.10.05', supplier: 'S', partNo: 'P', problem: '찍힘', typeManual: '기타', action: '재작업', typeIn: '외관' }, null, 3],
    types: [{ name: '내 유형', keywords: ['가', '가', ''], actions: [{ text: '대책', weeks: 0 }] }, { name: '미분류' }],
    rules: { repeatMin: 4 }, mask: false,
    plans: { S: { baseDate: '2026-10-26', rows: [{ id: 'x', action: 'a', status: '이상한값', planStart: '2026/11/2' }], confirmed: [{ at: '2026-10-20T10:00', rows: [] }] } }
  });
  assert.equal(db.records.length, 1);
  assert.equal(db.records[0].date, '2026-10-05');
  assert.equal(db.records[0].id, 'd1');
  assert.equal(db.records[0].action, '재작업'); assert.equal(db.records[0].typeIn, '외관');
  assert.deepEqual(db.types.map(t => t.name), ['내 유형']);
  assert.deepEqual(db.types[0].keywords, ['가']);
  assert.equal(db.types[0].actions[0].weeks, 1);
  assert.equal(db.rules.repeatMin, 4);
  assert.equal(db.mask, false);
  assert.equal(db.plans.S.rows[0].status, 'open');
  assert.equal(db.plans.S.rows[0].planStart, '2026-11-02');
  assert.equal(db.plans.S.confirmed[0].version, 1);
});

test('1단계(v1) 기본 사전으로 저장된 데이터는 8종으로 옮김 · 직접 고른 유형·계획 행도 따라감', () => {
  const V1 = ['이종 불량', '용접 균열', '용접 불량', '가공 치수불량', '도장 불량', '외관 불량', '조립·체결 불량', '누유·누수', '녹·부식'].map(name => ({ name, keywords: ['x'] }));
  const db = L.restoreDb({ schemaVersion: 1, types: V1, records: [{ date: '2026-10-05', supplier: 'S', problem: 'a', typeManual: '가공 치수불량' }],
    plans: { S: { rows: [{ id: 'r1', type: '용접 균열', action: 'a' }], confirmed: [{ at: '2026-10-20T10:00', rows: [{ id: 'r1', type: '조립·체결 불량', action: 'a' }] }] } } });
  assert.deepEqual(L.typeNames(db.types), L.typeNames(TYPES));
  assert.equal(db.records[0].typeManual, '위치/치수');
  assert.equal(db.plans.S.rows[0].type, '용접불량');
  assert.equal(db.plans.S.confirmed[0].rows[0].type, '미조임');
  // 사용자가 이름을 고친 사전은 그대로 둠
  const mine = L.restoreDb({ schemaVersion: 1, types: V1.slice(0, 8).concat([{ name: '내 유형', keywords: ['y'] }]) });
  assert.equal(mine.types[8].name, '내 유형');
});

console.log('직접 입력·엑셀 붙여넣기');
test('붙여넣기 — 머리행 없으면 「발생일자·협력사·품번·품명·불량유형·내용·수량·조치」 순서', () => {
  const t = '2026-10-05\t협력사 A\tTP-1\t커버\t용접 불량\t비드 불량\t3\t재작업\r\n2026.10.06\t협력사 B\t\t\t\t도장 흐름\t\t\n\n';
  const p = L.parsePaste(t);
  assert.equal(p.header, false); assert.equal(p.rows, 2);
  assert.equal(p.records.length, 2);
  assert.deepEqual(p.records[0], { date: '2026-10-05', partName: '커버', partNo: 'TP-1', problem: '비드 불량', supplier: '협력사 A', qty: 3, typeIn: '용접 불량', action: '재작업' });
  assert.equal(p.records[1].date, '2026-10-06');
  assert.equal(L.enrich(p.records, TYPES)[1].type, '기타');
});
test('붙여넣기 — 머리행이 있으면 열 이름으로(순서 상관없음) · 칸 안 줄바꿈("…") · 내용 빈 줄은 건너뜀', () => {
  const t = '조치\t불량유형\t업체\t발생일\t내용\n"1차 선별\n2차 반품"\t위치·치수\t협력사 D\t2026-09-01\t"홀 ""A"" 위치 틀림"\n\t\t협력사 D\t2026-09-02\t';
  const p = L.parsePaste(t);
  assert.equal(p.header, true);
  assert.equal(p.records.length, 1); assert.equal(p.skipped, 1);
  assert.equal(p.records[0].action, '1차 선별 2차 반품');
  assert.equal(p.records[0].problem, '홀 "A" 위치 틀림');
  assert.equal(L.enrich(p.records, TYPES)[0].type, '위치/치수');
});
test('붙여넣기 — 데이터 줄을 머리행으로 오해하지 않음(날짜가 든 줄)', () => {
  const p = L.parsePaste('2026-10-05\t협력사 내용\t\t\t\t유형 확인 필요');
  assert.equal(p.header, false);
  assert.equal(p.records[0].supplier, '협력사 내용');
});
test('직접 입력 점검 — 필수(발생일자·협력사·내용) · 수량 숫자 · 유형은 목록에서', () => {
  const ok = { date: '2026-10-05', supplier: '협력사 A', problem: '와셔 누락', type: '이종/누락', qty: '2', action: '선별' };
  assert.deepEqual(L.checkRecordInput(ok, TYPES), []);
  assert.equal(L.checkRecordInput({ ...ok, date: '' }, TYPES).length, 1);
  assert.equal(L.checkRecordInput({ ...ok, supplier: ' ', problem: '' }, TYPES).length, 2);
  assert.equal(L.checkRecordInput({ ...ok, qty: '두 개' }, TYPES).length, 1);
  assert.equal(L.checkRecordInput({ ...ok, type: '없는 유형' }, TYPES).length, 1);
  assert.equal(L.checkRecordInput({ ...ok, type: '' }, TYPES).length, 0); // 비우면 내용 키워드로 자동
  const r = L.recordFromInput(ok);
  assert.deepEqual(r, { date: '2026-10-05', partName: '', partNo: '', problem: '와셔 누락', supplier: '협력사 A', qty: 2, typeManual: '이종/누락', action: '선별' });
  assert.equal(L.enrich([r], TYPES)[0].manual, true);
});
test('엑셀 백업 → 다시 읽으면 같은 레코드(붙여넣기·파일 복원 모두)', () => {
  const recs = REC.slice(0, 20).concat([{ id: 'x', date: '2026-10-05', supplier: '협력사 Z', partNo: '', partName: '', problem: '밸브 작동 불량', typeManual: '성능이상', action: '교체', qty: 1 }]);
  const aoa = L.backupSheet(recs);
  assert.deepEqual(aoa[0], ['발생일자', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치']);
  const back = L.parseSheet(aoa).records;
  assert.equal(back.length, recs.length);
  assert.deepEqual(back.map(L.recordKey), recs.map(L.recordKey));
  assert.equal(L.enrich(back, TYPES)[20].type, '성능이상');
  assert.equal(back[20].action, '교체');
  const pasted = L.parsePaste(aoa.map(r => r.join('\t')).join('\n'));
  assert.equal(pasted.header, true); assert.equal(pasted.records.length, recs.length);
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
