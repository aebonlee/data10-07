// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 실제 공정불량 Raw data 는 아직 받지 않았습니다. 여기서는 가상 표와 예시 데이터(협력사 A~F)로 검사합니다.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const E = require('../js/ai-endpoint.js');
const XLSX = require('../vendor/xlsx.full.min.js');

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
  assert.equal(d.map.supplier, 2); assert.equal(d.map.partNo, 3); assert.equal(d.map.problem, 5); assert.equal(d.map.qty, 6);
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
test('키워드로 분류 · 맞은 키워드를 근거로 남김', () => {
  const c = L.classify('브래킷 용접부 갈라짐', TYPES);
  assert.equal(c.type, '용접 균열');
  assert.deepEqual(c.hits, ['갈라짐']);
  assert.equal(L.classify('이종품 혼입 (타 품번 섞임)', TYPES).type, '이종 불량');
  assert.equal(L.classify('홀 위치 치수 공차 초과', TYPES).type, '가공 치수불량');
  assert.equal(L.classify('도장 흐름 발생', TYPES).type, '도장 불량');
});
test('여러 유형이 걸리면 긴 키워드 → 같으면 사전에서 위에 있는 유형', () => {
  assert.equal(L.classify('용접 비드 균열', TYPES).type, '용접 균열'); // 비드·균열 모두 2글자 → 위의 용접 균열
  assert.equal(L.classify('용접 기공 다수', TYPES).type, '용접 불량');
  assert.equal(L.classify('공급사 언더컷 결함', TYPES).type, '용접 불량');
  assert.equal(L.classify('Crack 발생', TYPES).type, '용접 균열'); // 대소문자 무시
  const custom = [{ name: '외관', keywords: ['찍힘'] }, { name: '포장', keywords: ['포장 찍힘'] }];
  assert.equal(L.classify('포장 찍힘 발견', custom).type, '포장'); // 긴 키워드(띄어쓰기 빼고 4글자)가 이김
});
test('맞는 키워드가 없으면 미분류 · 직접 고른 유형이 먼저', () => {
  assert.equal(L.classify('검사성적서 미첨부', TYPES).type, L.UNCLASSIFIED);
  const r = { problem: '검사성적서 미첨부', typeManual: '외관 불량' };
  assert.deepEqual(L.typeOf(r, TYPES), { type: '외관 불량', hits: [], manual: true });
  assert.equal(L.typeOf({ problem: '용접 기공', typeManual: '사전에 없는 유형' }, TYPES).type, '용접 불량');
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
test('심어 둔 신호 — 협력사 C 용접 균열 증가 추세 · 반복 품번 TP-3301', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const i = an.issues.find(x => x.type === '용접 균열');
  assert.ok(i, '용접 균열이 개선 필요 항목에 있어야 함');
  assert.ok(i.flags.trend, '증가 추세');
  assert.equal(i.recent, 7); assert.equal(i.prev, 0);
  assert.ok(i.repeatParts.some(p => p.partNo === 'TP-3301'));
});
test('심어 둔 신호 — 협력사 A 이종 불량·TP-1102, 협력사 D 가공 치수불량·TP-4405', () => {
  const a = L.analyze(ALL, { supplier: '협력사 A' }, {}).issues.find(x => x.type === '이종 불량');
  assert.ok(a && a.flags.repeat && a.repeatParts[0].partNo === 'TP-1102');
  const d = L.analyze(ALL, { supplier: '협력사 D' }, {}).issues.find(x => x.type === '가공 치수불량');
  assert.ok(d && d.repeatParts[0].partNo === 'TP-4405');
});
test('협력사 F 는 건수가 적어 규칙에 걸리지 않음', () => {
  assert.equal(L.analyze(ALL, { supplier: '협력사 F' }, {}).issues.length, 0);
});
test('규칙 경계 — 반복 기준 딱 맞으면 걸리고 하나 모자라면 안 걸림 · 미분류는 항목에서 뺌', () => {
  const mk = (n, problem) => Array.from({ length: n }, (_, i) => ({ date: '2026-0' + (1 + i % 3) + '-1' + i, supplier: 'S', partNo: 'P' + i, problem }));
  const rules = { repeatMin: 3, partRepeatMin: 2, trendMonths: 3, trendRate: 50, trendMin: 9 };
  assert.equal(L.analyze(L.enrich(mk(3, '찍힘'), TYPES), {}, rules).issues.length, 1);
  assert.equal(L.analyze(L.enrich(mk(2, '찍힘'), TYPES), {}, rules).issues.length, 0);
  assert.equal(L.analyze(L.enrich(mk(5, '성적서 누락 안내'), TYPES.filter(t => t.name !== '조립·체결 불량')), {}, rules).issues.length, 0);
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
  const crack = rows.filter(r => r.type === '용접 균열');
  assert.equal(crack.length, TYPES.find(t => t.name === '용접 균열').actions.length);
  const k = types.indexOf('용접 균열');
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
  const plan = { rows: [L.makeRow({ supplier: 'S', type: '외관 불량', action: '포장 개선', planStart: '2026-11-02', planEnd: '2026-11-15' }, newId)] };
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
    '**요약:** 업체3 은 용접 균열이 최근 늘었고 ' + P3301 + ' 에서 반복돼.',
    '',
    '| 불량유형 | 개선대책 | 담당부서 | 시작(기준일부터 몇 주 뒤) | 기간(주) |',
    '|:---|---|---|---|---|',
    '| 용접균열 | ' + P3301 + ' 용접 예열 온도 기록 시작 | 협력사 | 0 | 2주 |',
    '| 용접 균열 | PT 검사<br>샘플링 2배 | 부품품질팀 | 1주 | 4 |',
    '| 새 유형 | 기타 대책 | | | |',
    '| 잘린 줄 |',
    '확인 필요:',
    '- 예열 설비가 있는지',
    '- 품목 3개 공정 순서'
  ].join('\n');
  const p = L.parseAiPlan(ans, { types: TYPES, aliases: AL, supplier: '협력사 C', baseDate: '2026-10-26', newId });
  assert.equal(p.summary, '협력사 C 은 용접 균열이 최근 늘었고 TP-3301 에서 반복돼.');
  assert.equal(p.rows.length, 3);
  assert.equal(p.skipped, 1);
  assert.equal(p.rows[0].type, '용접 균열');
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
  const p = L.parseAiPlan('도장 불량\t도막 측정 추가\t부품품질팀\t2\t3', { types: TYPES, aliases: {}, supplier: 'S', baseDate: '2026-10-26', newId });
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
test('보고서 시트 5개 — Master Plan 주 칸 표시(○ 계획 · ● 실적) · 상세 · 확정 이력', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 C', baseDate: '2026-10-26', newId });
  rows[0].actStart = '2026-10-26'; rows[0].actEnd = '2026-10-30'; rows[0].status = 'done';
  const plan = { baseDate: '2026-10-26', from: '', to: '', rows, confirmed: [] };
  L.confirmPlan(plan, '2026-10-20T11:00:00');
  const sh = L.reportSheets({ suppliers: ['협력사 C'], records: ALL, plans: { '협력사 C': plan }, rules: {}, today: '2026-10-20' });
  assert.deepEqual(Object.keys(sh), ['요약', 'Master Plan', '개선대책 상세', '불량 상세', '확정 이력']);
  const mp = sh['Master Plan'];
  assert.equal(mp[0][12], '2026년 10월 5주차');
  assert.equal(mp.length, rows.length + 3);
  assert.equal(mp[1][10], '완료');
  assert.equal(mp[1][12], '●');
  assert.equal(mp[2][12], '');   // 둘째 활동은 11월 2일 주부터
  assert.equal(mp[2][13], '○');
  assert.equal(sh['불량 상세'].length, an.total + 1);
  assert.equal(sh['개선대책 상세'].length, rows.length + 1);
  assert.deepEqual(sh['확정 이력'][1], ['협력사 C', 'v1', '2026-10-20 11:00', rows.length]);
  assert.ok(sh['요약'].some(r => r[0] === '계획 상태' && r[1].startsWith('확정 v1')));
});
test('보고서를 SheetJS 로 실제 xlsx 로 쓰고 다시 읽음', () => {
  const sh = L.reportSheets({ suppliers: ['협력사 A'], records: ALL, plans: {}, rules: {}, today: '2026-10-20' });
  const wb = XLSX.utils.book_new();
  Object.keys(sh).forEach(n => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh[n]), n));
  const back = XLSX.read(XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }), { type: 'buffer' });
  assert.deepEqual(back.SheetNames, ['요약', 'Master Plan', '개선대책 상세', '불량 상세', '확정 이력']);
});
test('저장 형식 — 깨진 값은 버리고 기본값 · 사용자 사전·계획 유지', () => {
  assert.equal(L.restoreDb(null).types.length, TYPES.length);
  const db = L.restoreDb({
    records: [{ date: '2026.10.05', supplier: 'S', partNo: 'P', problem: '찍힘', typeManual: '외관 불량' }, null, 3],
    types: [{ name: '내 유형', keywords: ['가', '가', ''], actions: [{ text: '대책', weeks: 0 }] }, { name: '미분류' }],
    rules: { repeatMin: 4 }, mask: false,
    plans: { S: { baseDate: '2026-10-26', rows: [{ id: 'x', action: 'a', status: '이상한값', planStart: '2026/11/2' }], confirmed: [{ at: '2026-10-20T10:00', rows: [] }] } }
  });
  assert.equal(db.records.length, 1);
  assert.equal(db.records[0].date, '2026-10-05');
  assert.equal(db.records[0].id, 'd1');
  assert.deepEqual(db.types.map(t => t.name), ['내 유형']);
  assert.deepEqual(db.types[0].keywords, ['가']);
  assert.equal(db.types[0].actions[0].weeks, 1);
  assert.equal(db.rules.repeatMin, 4);
  assert.equal(db.mask, false);
  assert.equal(db.plans.S.rows[0].status, 'open');
  assert.equal(db.plans.S.rows[0].planStart, '2026-11-02');
  assert.equal(db.plans.S.confirmed[0].version, 1);
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
