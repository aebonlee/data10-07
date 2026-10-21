/* 협력사 공정불량 분석 · 품질개선 Master Plan — 순수 로직 (화면·저장소와 무관, node 에서도 불러 씀)

   1. 날짜·주차          5. 개선 필요 항목(규칙)      9. AI 프롬프트·답 읽기(별칭 가리기)
   2. 엑셀 열 찾기·읽기   6. Master Plan 초안           10. 보고서(Excel 시트)
   3. 불량유형 분류       7. 간트 칸·추진 집계           11. 저장 형식
   4. 분석(건수·유형·월별·반복 품번)   8. 계획 확정
*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SQLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function str(v) { return v == null ? '' : String(v).trim(); }
  function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
  function round(n, d) { var p = Math.pow(10, d || 0); return Math.round(n * p) / p; }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  /* ── 1. 날짜·주차 ─────────────────────────────────────────── */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function validYmd(y, m, d) {
    if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
    var t = new Date(Date.UTC(y, m - 1, d));
    return t.getUTCMonth() === m - 1 ? ymd(y, m, d) : '';
  }
  /* 엑셀 날짜 일련번호 · Date · 「2026-10-05」「2026.10.5」「2026/10/05」「2026년 10월 5일」「26.10.05」「20261005」 */
  function parseDate(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) return isNaN(v) ? '' : validYmd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) { var t = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000); return ymd(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); }
      if (v >= 19900101 && v <= 21001231) return parseDate(String(v));
      return '';
    }
    var s = str(v), m;
    if ((m = /^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/.exec(s))) return validYmd(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{2})[-./](\d{1,2})[-./](\d{1,2})$/.exec(s))) return validYmd(2000 + +m[1], +m[2], +m[3]);
    if ((m = /^(\d{4})(\d{2})(\d{2})$/.exec(s))) return validYmd(+m[1], +m[2], +m[3]);
    return '';
  }
  function toDateStr(d) { return ymd(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function utc(s) { return Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)); }
  function fromUtc(t) { var d = new Date(t); return ymd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
  function addDays(s, n) { return fromUtc(utc(s) + n * 86400000); }
  function daysBetween(a, b) { return Math.round((utc(b) - utc(a)) / 86400000); }
  function dow(s) { return new Date(utc(s)).getUTCDay(); } // 0=일
  function mondayOf(s) { var w = dow(s); return addDays(s, w === 0 ? -6 : 1 - w); }
  function monthOf(s) { return s ? s.slice(0, 7) : ''; }
  function monthAdd(ym, k) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + k; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12; return y + '-' + pad(m + 1); }
  function monthLast(ym) { return addDays(monthAdd(ym, 1) + '-01', -1); }
  /* 주차 — 월~일 한 주는 「목요일이 든 달」의 주로 셉니다(2026-10-26 ~ 11-01 → 2026년 10월 5주차) */
  function weekInfo(s) {
    var mon = mondayOf(s), thu = addDays(mon, 3);
    var y = +thu.slice(0, 4), m = +thu.slice(5, 7), w = Math.floor((+thu.slice(8, 10) - 1) / 7) + 1;
    return { monday: mon, sunday: addDays(mon, 6), year: y, month: m, week: w, label: y + '년 ' + m + '월 ' + w + '주차' };
  }

  /* ── 2. 엑셀 열 찾기·읽기 ──────────────────────────────────── */
  var FIELDS = [
    { id: 'date', label: '발생일자', req: true, aliases: ['발생일자', '발생일', '발생날짜', '일자', '날짜', '접수일', '접수일자', '발견일', '발견일자', 'date'] },
    { id: 'partName', label: '품명', req: false, aliases: ['품명', '품목명', '부품명', '품목', 'partname'] },
    { id: 'partNo', label: '품번', req: false, aliases: ['품번', '부품번호', '품목번호', '품목코드', 'partno', 'p/n', 'pn'] },
    { id: 'problem', label: '내용(문제점)', req: true, aliases: ['문제점', '불량내용', '불량현상', '문제내용', '현상', '내용', '불량명', 'defect'] },
    { id: 'supplier', label: '협력사', req: true, aliases: ['협력사', '협력사명', '협력업체', '업체', '업체명', '공급사', '거래처', 'supplier', 'vendor'] },
    { id: 'type', label: '불량유형(선택)', req: false, aliases: ['불량유형', '유형', '불량구분', '불량분류', '불량코드', 'type', 'defecttype'] },
    { id: 'qty', label: '수량(선택)', req: false, aliases: ['수량', '불량수량', 'qty'] },
    { id: 'action', label: '조치(선택)', req: false, aliases: ['조치', '조치내용', '조치사항', '처리', '처리내용', '처리결과', 'action'] }
  ];
  /* 엑셀 붙여넣기에 머리행이 없을 때의 열 순서 = 불량 등록 화면·붙여넣기 양식의 순서 */
  var PASTE_ORDER = ['date', 'supplier', 'partNo', 'partName', 'type', 'problem', 'qty', 'action'];
  function fieldLabel(id) { var f = FIELDS.filter(function (x) { return x.id === id; })[0]; return f ? f.label.replace(/\(.*\)$/, '') : id; }
  function normHead(v) { return str(v).replace(/\(.*?\)|\[.*?\]/g, '').replace(/[\s·.:_\-]/g, '').toLowerCase(); }
  function matchField(h, taken) {
    var n = normHead(h), i, f;
    if (!n) return null;
    for (i = 0; i < FIELDS.length; i++) { f = FIELDS[i]; if (!taken[f.id] && f.aliases.indexOf(n) >= 0) return f.id; }
    for (i = 0; i < FIELDS.length; i++) {
      f = FIELDS[i];
      if (!taken[f.id] && f.aliases.some(function (a) { return a.length >= 2 && n.indexOf(a) === 0; })) return f.id;
    }
    return null;
  }
  /* 위 15행 안에서 머리행을 찾습니다(제목 줄이 있는 대장 형식도 읽기 위해) */
  function detectColumns(aoa) {
    var best = { headerRow: -1, map: {}, score: 0 };
    for (var r = 0; r < Math.min(15, aoa.length); r++) {
      var row = aoa[r] || [], map = {}, score = 0;
      for (var c = 0; c < row.length; c++) {
        var id = matchField(row[c], map);
        if (id) { map[id] = c; score++; }
      }
      if (score > best.score) best = { headerRow: r, map: map, score: score };
    }
    if (best.score < 2) best = { headerRow: -1, map: {}, score: 0 };
    var headers = best.headerRow >= 0 ? (aoa[best.headerRow] || []).map(str) : [];
    var missing = FIELDS.filter(function (f) { return f.req && best.map[f.id] == null; }).map(function (f) { return f.id; });
    return { headerRow: best.headerRow, map: best.map, headers: headers, missing: missing };
  }
  function parseRows(aoa, det) {
    var out = [], badDate = 0, skipped = 0;
    if (det.headerRow < 0 || det.missing.length) return { records: out, badDate: 0, skipped: 0 };
    var M = det.map;
    function cell(row, id) { return M[id] == null ? null : row[M[id]]; }
    for (var r = det.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var rec = {
        date: parseDate(cell(row, 'date')),
        partName: str(cell(row, 'partName')),
        partNo: str(cell(row, 'partNo')),
        problem: str(cell(row, 'problem')).replace(/\s*\n\s*/g, ' '),
        supplier: str(cell(row, 'supplier'))
      };
      var q = Number(cell(row, 'qty'));
      if (M.qty != null && isFinite(q) && str(cell(row, 'qty')) !== '') rec.qty = q;
      if (str(cell(row, 'type'))) rec.typeIn = str(cell(row, 'type'));
      if (str(cell(row, 'action'))) rec.action = str(cell(row, 'action')).replace(/\s*\n\s*/g, ' ');
      if (!rec.problem && !rec.supplier && !rec.partNo) { if (row.some(function (v) { return str(v); })) skipped++; continue; }
      // 내용이 비었어도 불량유형 칸이 있으면 그것을 내용으로 씁니다(유형만 적는 대장도 있어서)
      if (!rec.problem && rec.typeIn) rec.problem = rec.typeIn;
      if (!rec.problem || !rec.supplier) { skipped++; continue; }
      if (!rec.date) badDate++;
      out.push(rec);
    }
    return { records: out, badDate: badDate, skipped: skipped };
  }
  /* overrides: { 항목id: 열번호(-1 = 연결 끊기) } — 사용자가 열 연결을 고친 경우 */
  function parseSheet(aoa, opts) {
    var det = detectColumns(aoa);
    var ov = opts && opts.overrides;
    if (ov) {
      if (det.headerRow < 0 && opts.headerRow != null) { det.headerRow = opts.headerRow; det.headers = (aoa[opts.headerRow] || []).map(str); }
      Object.keys(ov).forEach(function (k) { if (ov[k] == null || ov[k] < 0) delete det.map[k]; else det.map[k] = ov[k]; });
      det.missing = FIELDS.filter(function (f) { return f.req && det.map[f.id] == null; }).map(function (f) { return f.id; });
    }
    var p = parseRows(aoa, det);
    return { det: det, records: p.records, badDate: p.badDate, skipped: p.skipped };
  }
  /* 여러 시트 중 머리행이 가장 잘 맞는 시트 */
  function pickSheet(sheets) {
    var best = null, bestScore = -1;
    sheets.forEach(function (s) {
      var d = detectColumns(s.aoa), sc = d.headerRow < 0 ? -1 : Object.keys(d.map).length * 10 - d.missing.length * 20;
      if (sc > bestScore) { bestScore = sc; best = s; }
    });
    return best;
  }
  function recordKey(r) { return [r.date, str(r.supplier).replace(/\s+/g, ''), str(r.partNo).replace(/\s+/g, '').toUpperCase(), str(r.problem).replace(/\s+/g, '')].join('|'); }
  /* 같은 행(발생일자+협력사+품번+문제점)은 한 번만. 같은 날 같은 불량이 정말 두 건이면 수량 열로 구분해 주세요 */
  function mergeRecords(old, add, seq) {
    var seen = {}, out = old.slice(), added = 0, dup = 0;
    seq = seq || 0;
    old.forEach(function (r) { seen[recordKey(r)] = 1; });
    add.forEach(function (r) {
      var k = recordKey(r);
      if (seen[k]) { dup++; return; }
      seen[k] = 1;
      var c = clone(r); c.id = 'd' + (++seq); out.push(c); added++;
    });
    return { records: out, added: added, dup: dup, seq: seq };
  }
  /* 엑셀에서 복사한 표(탭으로 나뉜 글) → 배열의 배열. 칸 안 줄바꿈은 엑셀이 "…" 로 감싸 줍니다 */
  function splitTsv(text) {
    var s = String(text == null ? '' : text).replace(/\r\n?/g, '\n'), rows = [], row = [], cell = '', i = 0, q = false;
    while (i < s.length) {
      var ch = s.charAt(i);
      if (q) {
        if (ch === '"' && s.charAt(i + 1) === '"') { cell += '"'; i += 2; continue; }
        if (ch === '"') { q = false; i++; continue; }
        cell += ch; i++; continue;
      }
      if (ch === '"' && cell === '') { q = true; i++; continue; }
      if (ch === '\t') { row.push(cell); cell = ''; i++; continue; }
      if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; i++; continue; }
      cell += ch; i++;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (v) { return str(v); }); });
  }
  /* 붙여넣기 읽기 — 머리행(열 이름)이 있으면 그 이름으로, 없으면 PASTE_ORDER 순서로 읽습니다.
     날짜처럼 읽히는 칸이 있는 줄은 머리행으로 보지 않습니다(데이터 줄을 머리행으로 오해하지 않게) */
  function parsePaste(text) {
    var aoa = splitTsv(text);
    if (!aoa.length) return { records: [], badDate: 0, skipped: 0, header: false, rows: 0, missing: [] };
    var det = detectColumns(aoa);
    var header = det.headerRow >= 0 && !(aoa[det.headerRow] || []).some(function (v) { return parseDate(v); });
    var p;
    if (header) p = parseSheet(aoa);
    else {
      var map = {}; PASTE_ORDER.forEach(function (id, c) { map[id] = c; });
      p = parseRows([PASTE_ORDER.slice()].concat(aoa), { headerRow: 0, map: map, missing: [] });
      p.det = { headerRow: -1, map: map, missing: [] };
    }
    return { records: p.records, badDate: p.badDate, skipped: p.skipped, header: header, rows: aoa.length - (header ? det.headerRow + 1 : 0), missing: p.det.missing || [] };
  }
  /* 화면에서 직접 등록하는 한 건 — 점검 후 저장용 레코드로 */
  function checkRecordInput(o, types) {
    var e = [];
    if (!str(o.date)) e.push('발생일자를 골라 주세요.');
    else if (!parseDate(o.date)) e.push('발생일자를 읽지 못했습니다(예: 2026-10-05).');
    if (!str(o.supplier)) e.push('협력사를 적어 주세요.');
    if (!str(o.problem)) e.push('내용(문제점)을 적어 주세요.');
    if (str(o.qty) !== '' && !(isFinite(Number(o.qty)) && Number(o.qty) >= 0)) e.push('수량은 0 이상의 숫자로 적어 주세요.');
    if (str(o.type) && str(o.type) !== UNCLASSIFIED && typeNames(types).indexOf(str(o.type)) < 0) e.push('불량유형을 목록에서 골라 주세요.');
    return e;
  }
  function recordFromInput(o) {
    var r = { date: parseDate(o.date), partName: str(o.partName), partNo: str(o.partNo), problem: str(o.problem).replace(/\s*\n\s*/g, ' '), supplier: str(o.supplier) };
    if (str(o.qty) !== '') r.qty = Number(o.qty);
    if (str(o.type)) r.typeManual = str(o.type);
    if (str(o.action)) r.action = str(o.action);
    return r;
  }

  /* ── 3. 불량유형 분류 (키워드 사전, 설정에서 고침) ─────────────── */
  var UNCLASSIFIED = '미분류';
  var DEFAULT_DEPTS = ['부품품질팀', '협력사', '구매팀', '생산기술팀', '설계팀'];
  /* actions: [대책, 담당부서, 시작(기준일부터 몇 주 뒤), 기간(주)] — 1단계 기본 대책(가정). 실제 코드표·부서를 받으면 바꿉니다 */
  var DEFAULT_TYPES = [
    /* 수강생이 정한 불량유형 8종(2026-10-21 답변). 키워드·기본 대책은 1단계 가정 — 설정에서 고칩니다.
       도장·외관·녹 같은 것은 8종에 없어 「기타」의 키워드로 두었습니다. */
    { name: '위치/치수', dept: '협력사', keywords: ['치수', '공차', '홀 위치', '홀위치', '위치 불량', '위치 틀림', '위치 상이', '피치', '길이', '외경', '내경', '두께', '높이', '편심', '평면도', '동심', '진원', '규격 초과', '규격 미달', '가공', '탭', '나사산', 'burr', '바리'],
      actions: [['치수·위치 불량 부위 측정, 공정능력(Cpk) 확인', '협력사', 0, 2], ['지그·기준면·공구 마모 점검, 가공 조건 재설정', '협력사', 1, 2], ['초·중·종물 측정 기록 의무화', '협력사', 2, 2], ['도면 공차·기준 위치 적정성 검토', '설계팀', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 5, 4]] },
    { name: '간섭/정렬', dept: '협력사', keywords: ['간섭', '정렬', '틀어짐', '어긋남', '단차', '유격', '센터', '직각', '평행', '비틀림', '닿음', '맞지 않음', '조립 불가', '조립불가', '쏠림', '오조립', '역조립', '방향 반대'],
      actions: [['간섭·정렬 불량 부위와 상대 부품 조사(조립 상태 재현)', '협력사', 0, 1], ['조립 지그·위치 결정핀 점검, 정렬 기준 한도견본 제작', '협력사', 1, 2], ['상대 부품 공차 누적 검토', '설계팀', 1, 2], ['조립 순서 작업표준서 개정', '생산기술팀', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 4, 4]] },
    { name: '이종/누락', dept: '협력사', keywords: ['이종', '오품', '혼입', '혼용', '타품', '오삽', '오출하', '품번 상이', '품번 오류', '라벨 오류', '식별 오류', '사양 상이', '누락', '결품', '빠짐', '미장착', '미부착', '미삽입', '수량 부족'],
      actions: [['이종·누락 발생 경로 조사(입고·보관·가공·조립·출하 단계별)', '협력사', 0, 1], ['품번별 식별표·보관 위치 구분, 부품 키팅·체크시트 도입', '협력사', 1, 2], ['출하 전 라벨·바코드 대조, 누락 확인 항목 추가', '협력사', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 4, 4]] },
    { name: '용접불량', dept: '협력사', keywords: ['용접', '균열', '크랙', 'crack', '갈라짐', '기공', '블로우홀', '언더컷', '스패터', '미용접', '비드', '각장', '용입', '오버랩'],
      actions: [['용접 결함 유형별 발생 위치·원인 분석(용접 조건·모재·예열)', '협력사', 0, 2], ['용접 조건(전류·전압·속도·예열) 표준 재설정, 작업표준서 개정', '협력사', 2, 2], ['용접부 외관·비파괴검사(PT/MT) 샘플링 강화', '부품품질팀', 2, 4], ['용접사 교육·자격 점검, 지그·용접 순서 점검', '협력사', 3, 1], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 6, 4]] },
    { name: '누유/누수', dept: '협력사', keywords: ['누유', '누수', '리크', 'leak', '오일 샘', '물 샘', '기밀', '새어', '맺힘'],
      actions: [['기밀 시험 조건(압력·시간) 점검', '협력사', 0, 2], ['실링면 가공 상태·오링·피팅 체결 확인', '협력사', 1, 2], ['누유/누수 검사 전수화 여부 검토', '부품품질팀', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 4, 4]] },
    { name: '미조임', dept: '협력사', keywords: ['미조임', '조임 불량', '조임불량', '미체결', '체결', '토크', '풀림', '헐거', '볼트', '너트'],
      actions: [['체결 토크 관리 기준·토크 공구 교정 점검', '협력사', 0, 2], ['체결 확인 마킹(토크 마킹)·체크시트 도입', '협력사', 1, 2], ['체결 순서 작업표준서 개정', '생산기술팀', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 4, 4]] },
    { name: '성능이상', dept: '협력사', keywords: ['성능', '작동', '동작', '기능', '소음', '이음', '진동', '발열', '압력 저하', '압력 부족', '출력', '응답', '고장', '불능'],
      actions: [['성능이상 재현 시험·원인 부품 분리', '협력사', 0, 2], ['출하 성능검사 항목·판정 기준 점검', '부품품질팀', 1, 2], ['설계 사양·시험 조건 검토', '설계팀', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 4, 4]] },
    { name: '기타', dept: '협력사', keywords: ['도장', '도막', '도막 두께', '페인트', '찍힘', '스크래치', '긁힘', '이물', '오염', '녹 발생', '녹발생', '부식', '변형', '외관', '포장'],
      actions: [['발생 내용 확인·원인 조사', '협력사', 0, 2], ['재발 방지 대책 수립·작업표준 반영', '협력사', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '부품품질팀', 4, 4]] }
  ];
  function defaultTypes() {
    return DEFAULT_TYPES.map(function (t) {
      return { name: t.name, dept: t.dept, keywords: t.keywords.slice(), actions: t.actions.map(function (a) { return { text: a[0], dept: a[1], start: a[2], weeks: a[3] }; }) };
    });
  }
  function normText(s) { return str(s).toLowerCase().replace(/\s+/g, ''); }
  /* 점수 = 맞은 키워드 가운데 가장 긴 것의 글자 수(긴 키워드가 더 구체적).
     같으면 사전에서 위에 있는 유형 — 그래서 사전 순서가 곧 우선순위입니다(「이종 볼트 혼용」 → 혼용·볼트 모두 2글자 → 위에 있는 이종/누락). */
  function classify(text, types) {
    var t = normText(text), best = null;
    (types || []).forEach(function (ty, i) {
      var hits = [], score = 0;
      (ty.keywords || []).forEach(function (k) {
        var n = normText(k);
        if (n && t.indexOf(n) >= 0 && hits.indexOf(k) < 0) { hits.push(k); score = Math.max(score, n.length); }
      });
      if (score > 0 && (!best || score > best.score)) best = { type: ty.name, hits: hits, score: score, index: i };
    });
    return best || { type: UNCLASSIFIED, hits: [], score: 0, index: -1 };
  }
  function typeNames(types) { return (types || []).map(function (t) { return t.name; }); }
  /* 유형 이름 맞추기 — 띄어쓰기·「/」「·」 같은 구분 기호는 무시(「위치·치수」「위치 / 치수」 → 위치/치수, 「용접 불량」 → 용접불량) */
  function normType(s) { return str(s).toLowerCase().replace(/[\s\/·・,.\-_()\[\]]/g, ''); }
  function findTypeName(s, types) {
    var n = normType(s), names = typeNames(types), i;
    if (!n) return '';
    for (i = 0; i < names.length; i++) if (normType(names[i]) === n) return names[i];
    if (n.length >= 2) for (i = 0; i < names.length; i++) { var m = normType(names[i]); if (m.length >= 2 && (n.indexOf(m) >= 0 || m.indexOf(n) >= 0)) return names[i]; }
    return '';
  }
  /* 순서: ① 사람이 고른 유형(typeManual) ② 엑셀의 불량유형 칸(typeIn)이 사전 이름과 맞으면 그것 ③ 내용 키워드 */
  function typeOf(rec, types) {
    if (rec.typeManual && (rec.typeManual === UNCLASSIFIED || typeNames(types).indexOf(rec.typeManual) >= 0)) return { type: rec.typeManual, hits: [], manual: true };
    var col = rec.typeIn ? findTypeName(rec.typeIn, types) : '';
    if (col) return { type: col, hits: [], manual: false, column: true };
    var c = classify(rec.problem + (rec.typeIn ? ' ' + rec.typeIn : ''), types);
    return { type: c.type, hits: c.hits, manual: false };
  }
  function enrich(records, types) {
    return records.map(function (r) {
      var t = typeOf(r, types), c = {};
      Object.keys(r).forEach(function (k) { c[k] = r[k]; });
      c.type = t.type; c.hits = t.hits; c.manual = t.manual; c.column = !!t.column;
      return c;
    });
  }
  /* 설정 화면의 「대책 | 부서 | 시작주 | 기간주」 줄 ↔ actions */
  function parseActionLines(text, defDept) {
    return String(text || '').split(/\r?\n/).map(function (line) {
      var p = line.split('|').map(str);
      if (!p[0]) return null;
      var s = parseInt(p[2], 10), w = parseInt(p[3], 10);
      return { text: p[0], dept: p[1] || defDept || '부품품질팀', start: isFinite(s) && s >= 0 ? s : 0, weeks: isFinite(w) && w > 0 ? w : 2 };
    }).filter(Boolean);
  }
  function formatActionLines(actions) {
    return (actions || []).map(function (a) { return a.text + ' | ' + a.dept + ' | ' + a.start + ' | ' + a.weeks; }).join('\n');
  }
  function parseKeywords(text) { return uniq(String(text || '').split(/[,\n،、]/).map(str).filter(Boolean)); }

  /* ── 4. 분석 ─────────────────────────────────────────────── */
  var DEFAULT_RULES = { repeatMin: 5, partRepeatMin: 3, trendMonths: 3, trendRate: 50, trendMin: 3 };
  function cleanRules(r) {
    var o = clone(DEFAULT_RULES);
    if (!r || typeof r !== 'object') return o;
    function n(k, min, max) { var v = Number(r[k]); if (isFinite(v) && v >= min && v <= max) o[k] = Math.round(v); }
    n('repeatMin', 1, 999); n('partRepeatMin', 2, 999); n('trendMonths', 1, 12); n('trendRate', 0, 1000); n('trendMin', 1, 999);
    return o;
  }
  function inRange(r, from, to) { return (!from || (r.date && r.date >= from)) && (!to || (r.date && r.date <= to)); }
  function countBy(rows, f) {
    var m = {};
    rows.forEach(function (r) { var k = f(r); m[k] = (m[k] || 0) + 1; });
    return Object.keys(m).map(function (k) { return { key: k, count: m[k] }; }).sort(function (a, b) { return b.count - a.count || (a.key < b.key ? -1 : 1); });
  }
  function dateBounds(rows) {
    var ds = rows.map(function (r) { return r.date; }).filter(Boolean).sort();
    return { min: ds[0] || '', max: ds[ds.length - 1] || '' };
  }
  /* rows 는 enrich() 를 거친 것. opts = { supplier, from, to } */
  function analyze(rows, opts, rules) {
    opts = opts || {}; rules = cleanRules(rules);
    var sel = rows.filter(function (r) { return (!opts.supplier || r.supplier === opts.supplier) && inRange(r, opts.from, opts.to); });
    var b = dateBounds(sel);
    var from = opts.from || b.min, to = opts.to || b.max;
    var total = sel.length;
    var byType = countBy(sel, function (r) { return r.type; });
    // 파레토: 미분류는 맨 뒤
    byType.sort(function (x, y) { return (x.key === UNCLASSIFIED) - (y.key === UNCLASSIFIED) || y.count - x.count || (x.key < y.key ? -1 : 1); });
    var cum = 0;
    byType = byType.map(function (x) { cum += x.count; return { type: x.key, count: x.count, pct: total ? round(x.count / total * 100, 1) : 0, cum: total ? round(cum / total * 100, 1) : 0 }; });
    var suppliers = countBy(sel, function (r) { return r.supplier; }).map(function (x) { return { name: x.key, count: x.count }; });
    // 월별 — 빈 달도 0 으로
    var months = [];
    if (from && to) {
      var mc = {};
      sel.forEach(function (r) { if (r.date) mc[monthOf(r.date)] = (mc[monthOf(r.date)] || 0) + 1; });
      for (var ym = monthOf(from), guard = 0; ym <= monthOf(to) && guard < 240; ym = monthAdd(ym, 1), guard++) months.push({ month: ym, count: mc[ym] || 0 });
    }
    // 반복 품번 — 같은 협력사·같은 품번
    var pm = {};
    sel.forEach(function (r) {
      if (!r.partNo) return;
      var k = r.supplier + '\u0000' + r.partNo;
      var g = pm[k] || (pm[k] = { supplier: r.supplier, partNo: r.partNo, partName: r.partName, count: 0, types: [], last: '' });
      g.count++; if (g.types.indexOf(r.type) < 0) g.types.push(r.type); if (r.date > g.last) g.last = r.date; if (!g.partName) g.partName = r.partName;
    });
    var repeatParts = Object.keys(pm).map(function (k) { return pm[k]; }).filter(function (g) { return g.count >= rules.partRepeatMin; })
      .sort(function (a, b) { return b.count - a.count || (a.partNo < b.partNo ? -1 : 1); });
    var issues = findIssues(sel, to, rules);
    return {
      from: from, to: to, rows: sel, total: total,
      qty: sel.reduce(function (s, r) { return s + (typeof r.qty === 'number' ? r.qty : 0); }, 0),
      hasQty: sel.some(function (r) { return typeof r.qty === 'number'; }),
      byType: byType, suppliers: suppliers, months: months, repeatParts: repeatParts, issues: issues,
      unclassified: sel.filter(function (r) { return r.type === UNCLASSIFIED; }).length
    };
  }

  /* ── 5. 개선 필요 항목 — 협력사×불량유형마다 규칙 3가지 ─────────────
     반복 발생: 기간 안 건수 ≥ repeatMin
     증가 추세: 최근 trendMonths 개월 건수 ≥ trendMin 이고, 직전 같은 길이 기간보다 trendRate% 이상 많음
     반복 품번: 같은 품번이 partRepeatMin 건 이상 */
  function findIssues(sel, to, rules) {
    rules = cleanRules(rules);
    var endM = monthOf(to || dateBounds(sel).max);
    if (!endM) return [];
    var n = rules.trendMonths, recentFrom = monthAdd(endM, -(n - 1)), prevFrom = monthAdd(endM, -(2 * n - 1)), prevTo = monthAdd(endM, -n);
    var g = {};
    sel.forEach(function (r) {
      if (r.type === UNCLASSIFIED) return;
      var k = r.supplier + '\u0000' + r.type;
      var x = g[k] || (g[k] = { supplier: r.supplier, type: r.type, count: 0, recent: 0, prev: 0, parts: {}, problems: [] });
      x.count++;
      var m = monthOf(r.date);
      if (m && m >= recentFrom && m <= endM) x.recent++;
      else if (m && m >= prevFrom && m <= prevTo) x.prev++;
      if (r.partNo) x.parts[r.partNo] = (x.parts[r.partNo] || 0) + 1;
      if (r.problem && x.problems.indexOf(r.problem) < 0) x.problems.push(r.problem);
    });
    var out = [];
    Object.keys(g).forEach(function (k) {
      var x = g[k], reasons = [], flags = {};
      if (x.count >= rules.repeatMin) { flags.repeat = true; reasons.push('반복 발생 ' + x.count + '건 (기준 ' + rules.repeatMin + '건 이상)'); }
      if (x.recent >= rules.trendMin && x.recent > x.prev && x.recent >= x.prev * (1 + rules.trendRate / 100)) {
        flags.trend = true;
        reasons.push('증가 추세 — 최근 ' + n + '개월 ' + x.recent + '건, 직전 ' + n + '개월 ' + x.prev + '건');
      }
      var rp = Object.keys(x.parts).filter(function (p) { return x.parts[p] >= rules.partRepeatMin; }).sort(function (a, b) { return x.parts[b] - x.parts[a] || (a < b ? -1 : 1); });
      if (rp.length) { flags.part = true; reasons.push('반복 품번 ' + rp.map(function (p) { return p + ' ' + x.parts[p] + '건'; }).join(', ')); }
      if (!reasons.length) return;
      out.push({ supplier: x.supplier, type: x.type, count: x.count, recent: x.recent, prev: x.prev, flags: flags, reasons: reasons,
        repeatParts: rp.map(function (p) { return { partNo: p, count: x.parts[p] }; }), problems: x.problems.slice(0, 5),
        score: x.count + (flags.trend ? 3 : 0) + rp.length });
    });
    return out.sort(function (a, b) { return b.score - a.score || b.count - a.count || (a.supplier + a.type < b.supplier + b.type ? -1 : 1); });
  }

  /* ── 6. Master Plan 초안 ─────────────────────────────────────
     개선 필요 항목의 불량유형마다 사전의 기본 대책을 넣습니다. 항목이 없으면 건수 상위 3개 유형.
     유형마다 1주씩 늦춰 시작해 한꺼번에 몰리지 않게 합니다. */
  var STATUS = { open: '미완료', done: '완료', hold: '보류' };
  var idCounter = 0;
  function defaultId() { idCounter++; return 'p' + Date.now().toString(36) + idCounter.toString(36); }
  function draftTypes(an, supplier, maxTypes) {
    var list = an.issues.filter(function (i) { return i.supplier === supplier; }).map(function (i) { return { type: i.type, reason: i.reasons.join(' / '), count: i.count }; });
    if (!list.length) {
      var cnt = countBy(an.rows.filter(function (r) { return r.supplier === supplier && r.type !== UNCLASSIFIED; }), function (r) { return r.type; });
      list = cnt.slice(0, 3).map(function (c) { return { type: c.key, reason: '건수 상위 유형 (' + c.count + '건, 규칙 기준 미만)', count: c.count }; });
    }
    return list.slice(0, maxTypes || 3);
  }
  function makeRow(o, newId) {
    return { id: (newId || defaultId)(), supplier: o.supplier, type: o.type, action: o.action, dept: o.dept || '부품품질팀', owner: o.owner || '',
      planStart: o.planStart || '', planEnd: o.planEnd || '', actStart: '', actEnd: '', status: 'open', note: o.note || '', source: o.source || 'manual' };
  }
  function buildDraft(an, types, opts) {
    var supplier = opts.supplier, base = opts.baseDate, out = [];
    var byName = {}; (types || []).forEach(function (t) { byName[t.name] = t; });
    draftTypes(an, supplier, opts.maxTypes).forEach(function (d, i) {
      var t = byName[d.type];
      var acts = t && t.actions && t.actions.length ? t.actions : [{ text: d.type + ' 원인 조사와 개선대책 수립', dept: '협력사', start: 0, weeks: 2 }, { text: '유효성 확인(대책 후 재발 여부 추적)', dept: '부품품질팀', start: 2, weeks: 4 }];
      acts.forEach(function (a) {
        var ps = addDays(base, 7 * (a.start + i));
        out.push(makeRow({ supplier: supplier, type: d.type, action: a.text, dept: a.dept, planStart: ps, planEnd: addDays(ps, 7 * a.weeks - 1), note: d.reason, source: 'template' }, opts.newId));
      });
    });
    return out;
  }

  /* ── 7. 간트 칸·추진 집계 ────────────────────────────────────── */
  function rowDates(r) { return [r.planStart, r.planEnd, r.actStart, r.actEnd].filter(function (d) { return /^\d{4}-\d\d-\d\d$/.test(d || ''); }); }
  function isOverdue(r, today) { return r.status === 'open' && !!r.planEnd && r.planEnd < today; }
  function planStats(rows, today) {
    var s = { total: rows.length, done: 0, open: 0, hold: 0, overdue: 0, start: '', end: '', days: 0, rate: 0 };
    rows.forEach(function (r) {
      s[r.status in STATUS ? r.status : 'open']++;
      if (isOverdue(r, today)) s.overdue++;
      if (r.planStart && (!s.start || r.planStart < s.start)) s.start = r.planStart;
      if (r.planEnd && (!s.end || r.planEnd > s.end)) s.end = r.planEnd;
    });
    if (s.start && s.end && s.end >= s.start) s.days = daysBetween(s.start, s.end) + 1;
    s.rate = s.total ? Math.round(s.done / s.total * 100) : 0;
    return s;
  }
  /* scale = 'week'(주차 칸, 위 줄 = 월) | 'month'(월 칸, 위 줄 = 연도) */
  function ganttColumns(rows, scale, opts) {
    opts = opts || {};
    var ds = [];
    rows.forEach(function (r) { ds = ds.concat(rowDates(r)); });
    if (opts.today) ds.push(opts.today);
    ds.sort();
    var first = ds[0] || opts.baseDate || opts.today, last = ds[ds.length - 1] || first;
    if (!first) return [];
    var cols = [], min = opts.minCols || (scale === 'month' ? 6 : 10);
    if (scale === 'month') {
      for (var m = monthOf(first); cols.length < 120 && (m <= monthOf(last) || cols.length < min); m = monthAdd(m, 1)) {
        cols.push({ start: m + '-01', end: monthLast(m), label: +m.slice(5) + '월', group: m.slice(0, 4) + '년', full: m.slice(0, 4) + '년 ' + (+m.slice(5)) + '월' });
      }
    } else {
      for (var w = mondayOf(first); cols.length < 260 && (w <= last || cols.length < min); w = addDays(w, 7)) {
        var wi = weekInfo(w);
        cols.push({ start: w, end: wi.sunday, label: wi.week + '주', group: wi.year + '년 ' + wi.month + '월', full: wi.label });
      }
    }
    return cols;
  }
  /* 기간이 걸치는 칸 번호 [처음, 끝] — 없으면 null. 끝이 비면 시작 칸 하나 */
  function span(start, end, cols) {
    if (!start || !cols.length) return null;
    end = end && end >= start ? end : start;
    var a = -1, b = -1;
    for (var i = 0; i < cols.length; i++) {
      if (a < 0 && cols[i].end >= start) a = i;
      if (cols[i].start <= end) b = i;
    }
    if (a < 0 || b < 0 || b < a) return null;
    return [a, b];
  }
  function groupSpans(cols) {
    var out = [];
    cols.forEach(function (c) { var g = out[out.length - 1]; if (g && g.group === c.group) g.n++; else out.push({ group: c.group, n: 1 }); });
    return out;
  }
  function cleanPlanRow(r) {
    var o = makeRow({ supplier: str(r.supplier), type: str(r.type), action: str(r.action), dept: str(r.dept), owner: str(r.owner), note: str(r.note), source: r.source === 'template' || r.source === 'ai' ? r.source : 'manual' }, function () { return str(r.id) || defaultId(); });
    ['planStart', 'planEnd', 'actStart', 'actEnd'].forEach(function (k) { o[k] = parseDate(r[k]); });
    o.status = r.status in STATUS ? r.status : 'open';
    return o;
  }
  /* 입력 점검 — 화면의 오른쪽 칸에서 저장 전에 */
  function checkPlanRow(r) {
    var e = [];
    if (!str(r.action)) e.push('개선대책을 적어 주세요.');
    if (r.planStart && r.planEnd && r.planEnd < r.planStart) e.push('계획 완료일이 시작일보다 빠릅니다.');
    if (r.actStart && r.actEnd && r.actEnd < r.actStart) e.push('실적 완료일이 시작일보다 빠릅니다.');
    if (r.status === 'done' && !r.actEnd) e.push('완료로 두려면 실적 완료일을 적어 주세요.');
    return e;
  }

  /* ── 8. 계획 확정 ──────────────────────────────────────────── */
  function planKey(rows) { return JSON.stringify(rows.map(function (r) { return [r.type, r.action, r.dept, r.owner, r.planStart, r.planEnd]; })); }
  function confirmPlan(plan, at) {
    plan.confirmed = plan.confirmed || [];
    var snap = { version: plan.confirmed.length + 1, at: at, rows: clone(plan.rows) };
    plan.confirmed.push(snap);
    return snap;
  }
  function lastConfirmed(plan) { return plan && plan.confirmed && plan.confirmed.length ? plan.confirmed[plan.confirmed.length - 1] : null; }
  /* 확정 뒤 계획(대책·부서·담당자·계획일)이 바뀌었는지 — 실적·상태 입력은 바뀜으로 보지 않습니다 */
  function planChangedSince(plan) { var c = lastConfirmed(plan); return !!c && planKey(c.rows) !== planKey(plan.rows); }
  function planStateText(plan) {
    var c = lastConfirmed(plan);
    if (!c) return '미확정 초안';
    return '확정 v' + c.version + ' (' + c.at.replace('T', ' ').slice(0, 16) + ')' + (planChangedSince(plan) ? ' 이후 계획 수정 있음' : '');
  }

  /* ── 9. AI 프롬프트·답 읽기 ─────────────────────────────────
     협력사 이름은 「업체1」, 품번은 「품목1」 같은 별칭으로 바꿉니다(기본). 대응표는 이 브라우저에만 둡니다. */
  var ALIAS = { suppliers: '업체', parts: '품목' };
  function assignAliases(map, names, prefix) {
    var out = {}, used = 0, re = new RegExp('^' + prefix + '(\\d+)$');
    Object.keys(map || {}).forEach(function (k) { var m = re.exec(map[k]); if (m) { out[k] = map[k]; used = Math.max(used, +m[1]); } });
    uniq((names || []).map(str).filter(Boolean)).forEach(function (k) { if (!out[k]) out[k] = prefix + (++used); });
    return out;
  }
  function updateAliases(aliases, records) {
    aliases = aliases || {};
    return {
      suppliers: assignAliases(aliases.suppliers, records.map(function (r) { return r.supplier; }).sort(), ALIAS.suppliers),
      parts: assignAliases(aliases.parts, records.map(function (r) { return r.partNo; }).sort(), ALIAS.parts)
    };
  }
  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&'); }
  /* 글 속에서 바꿀 수 있는 이름 — 두 글자 이상(한 글자는 다른 낱말과 헷갈림) */
  function nameRegex(names) {
    var list = names.filter(function (n) { return n.length >= 2; }).sort(function (a, b) { return b.length - a.length; }).map(escRe);
    return list.length ? new RegExp('(^|[^0-9A-Za-z])(' + list.join('|') + ')(?![0-9A-Za-z])', 'g') : null;
  }
  function maskText(text, aliases) {
    var s = text == null ? '' : String(text);
    var all = {};
    ['suppliers', 'parts'].forEach(function (k) { var m = (aliases && aliases[k]) || {}; Object.keys(m).forEach(function (n) { all[n] = m[n]; }); });
    var re = nameRegex(Object.keys(all));
    return re ? s.replace(re, function (x, pre, n) { return pre + all[n]; }) : s;
  }
  /* AI 답 → 원래 이름. 「업체3」「품목3」만 바꾸고 「품목 3개」「품목3개」처럼 수를 세는 말은 그대로 둡니다 */
  function unmaskText(text, aliases) {
    var back = {};
    ['suppliers', 'parts'].forEach(function (k) { var m = (aliases && aliases[k]) || {}; Object.keys(m).forEach(function (n) { back[m[n]] = n; }); });
    return String(text == null ? '' : text).replace(new RegExp('(' + ALIAS.suppliers + '|' + ALIAS.parts + ')(\\d+)(?![0-9개건종가])', 'g'), function (all) { return back[all] || all; });
  }
  function findLeaks(text, aliases) {
    var s = String(text || ''), names = [];
    ['suppliers', 'parts'].forEach(function (k) { names = names.concat(Object.keys((aliases && aliases[k]) || {})); });
    return names.filter(function (n) { var re = nameRegex([n]); return re && re.test(s); });
  }
  /* ctx = { supplier, an(analyze 결과, 협력사로 거른 것), rows(지금 계획), types, depts, baseDate, aliases(null 이면 가리지 않음) } */
  function buildAiPromptInfo(ctx) {
    var al = ctx.aliases || null, an = ctx.an, L = [];
    function M(s) { return al ? maskText(s, al) : String(s == null ? '' : s); }
    L.push('너는 부품품질팀의 협력사 품질개선 담당을 돕는 보조야. 아래 [사실]은 협력사 공정불량 Raw data 에서 도구가 계산한 값이야.');
    L.push('[사실]에 있는 숫자와 내용만 근거로 삼아줘. 숫자를 새로 만들거나 바꾸지 말고, 기록에 없는 원인은 반드시 「추정」이라고 표시해줘.');
    if (al) L.push('협력사 이름과 품번은 보안상 「' + ALIAS.suppliers + '1」「' + ALIAS.parts + '1」 같은 별칭으로 바꿔 두었어. 답에서도 별칭을 띄어 쓰지 말고 그대로 써줘(실제 이름을 짐작하지 말아줘).');
    L.push('');
    L.push('[사실]');
    L.push('- 협력사: ' + M(ctx.supplier));
    L.push('- 분석 기간: ' + an.from + ' ~ ' + an.to + ', 공정불량 ' + an.total + '건' + (an.unclassified ? ' (유형 미분류 ' + an.unclassified + '건 포함)' : ''));
    L.push('- 불량유형별 건수: ' + (an.byType.map(function (t) { return t.type + ' ' + t.count + '건(' + t.pct + '%)'; }).join(', ') || '없음'));
    var iss = an.issues.filter(function (i) { return i.supplier === ctx.supplier; });
    L.push('- 규칙으로 찾은 개선 필요 항목 ' + iss.length + '건');
    iss.forEach(function (i, n) { L.push('  ' + (n + 1) + '. ' + i.type + ' ' + i.count + '건 — ' + M(i.reasons.join('; '))); });
    var rp = an.repeatParts.filter(function (p) { return p.supplier === ctx.supplier; }).slice(0, 10);
    if (rp.length) L.push('- 반복 품번: ' + rp.map(function (p) { return M(p.partNo) + (p.partName ? '(' + p.partName + ')' : '') + ' ' + p.count + '건 [' + p.types.join(', ') + ']'; }).join('; '));
    var byType = {};
    an.rows.forEach(function (r) { if (r.supplier !== ctx.supplier) return; var a = byType[r.type] || (byType[r.type] = []); if (a.length < 4 && a.indexOf(r.problem) < 0) a.push(r.problem); });
    var tl = Object.keys(byType);
    if (tl.length) {
      L.push('- 유형별 문제점 기록(원문 일부):');
      tl.forEach(function (t) { L.push('  · ' + t + ': ' + byType[t].map(function (p) { return '「' + M(p) + '」'; }).join(' ')); });
    }
    L.push('- 담당부서 목록: ' + (ctx.depts || DEFAULT_DEPTS).join(', '));
    L.push('- 계획 착수 기준일: ' + ctx.baseDate);
    if (ctx.rows && ctx.rows.length) {
      L.push('- 지금 초안(기본 대책 사전에서 만든 것):');
      ctx.rows.forEach(function (r) { L.push('  · ' + r.type + ' | ' + M(r.action) + ' | ' + r.dept + ' | ' + r.planStart + '~' + r.planEnd); });
    }
    L.push('');
    L.push('[요청]');
    L.push('1. 첫 줄에 「요약:」으로 시작하는 한 줄 요약을 써줘.');
    L.push('2. 불량유형마다 구체적인 개선대책 2~4개를 아래 표 형식으로만 써줘. 지금 초안보다 이 협력사 기록에 맞게 구체적으로 써줘. 담당부서는 위 목록에서 골라줘.');
    L.push('| 불량유형 | 개선대책 | 담당부서 | 시작(기준일부터 몇 주 뒤) | 기간(주) |');
    L.push('|---|---|---|---|---|');
    L.push('3. 표 아래에 「확인 필요:」로 시작해 현장에서 확인할 것을 한 줄씩 적어줘.');
    var text = L.join('\n');
    return { prompt: text, leaks: al ? findLeaks(text, al) : [] };
  }
  function matchType(s, types) {
    if (!str(s)) return '';
    return findTypeName(s, types) || str(s);
  }
  function weeksNum(s, def) { var m = /(\d+)/.exec(str(s)); return m ? +m[1] : def; }
  /* AI 답 읽기 — 「| 유형 | 대책 | 부서 | 시작 | 기간 |」 표(또는 탭으로 나뉜 줄). opts = { types, aliases, supplier, baseDate, newId } */
  function parseAiPlan(text, opts) {
    var t = unmaskText(text, opts.aliases), rows = [], skipped = 0, summary = '', checks = [], inChecks = false;
    t.split(/\r?\n/).forEach(function (line) {
      var s = line.trim();
      if (!s) return;
      var sm = /^[#\s]*\**\s*요약\s*\**\s*[:：]\s*\**\s*(.+)$/.exec(s);
      if (sm && !summary) { summary = sm[1].replace(/\*+$/, '').trim(); return; }
      var cm = /^[#\s]*\**\s*확인\s*필요\s*\**\s*[:：]\s*\**\s*(.*)$/.exec(s);
      if (cm) { inChecks = true; if (cm[1].trim()) checks.push(cm[1].trim()); return; }
      var cells = null;
      if (s.charAt(0) === '|') cells = s.replace(/^\|/, '').replace(/\|$/, '').split('|').map(str);
      else if (s.indexOf('\t') >= 0) cells = s.split('\t').map(str);
      if (!cells) { if (inChecks) checks.push(s.replace(/^[-*·\d.)\s]+/, '')); return; }
      inChecks = false;
      if (cells.every(function (c) { return /^:?-{2,}:?$/.test(c) || !c; })) return;
      if (/불량\s*유형/.test(cells[0]) && /대책/.test(cells[1] || '')) return;
      if (cells.length < 2 || !cells[1]) { skipped++; return; }
      var st = weeksNum(cells[3], 0), wk = Math.max(1, weeksNum(cells[4], 2));
      var ps = addDays(opts.baseDate, 7 * st);
      rows.push(makeRow({ supplier: opts.supplier, type: matchType(cells[0], opts.types) || '기타', action: cells[1].replace(/<br\s*\/?>/gi, ' '), dept: cells[2] || '부품품질팀',
        planStart: ps, planEnd: addDays(ps, 7 * wk - 1), source: 'ai' }, opts.newId));
    });
    return { summary: summary, rows: rows, skipped: skipped, checks: checks };
  }

  /* ── 10. 보고서 (Excel 시트 = 배열의 배열 + 서식 표시) ─────────────
     웹 화면의 표 모양 그대로 6시트: 요약 · 불량 목록 · 유형별·협력사별 집계 · Master Plan · 간트 일정 · 확정 이력.
     meta 는 js/xlsx-out.js 가 머리글 굵게·칸 색·틀 고정·거르기를 붙일 때 씁니다.
     ctx = { suppliers:[이름], records(enrich 된 전체), plans, rules, today, types } */
  var REPORT_TITLE = '협력사 공정불량 분석 및 품질개선 Master Plan 보고서';
  function sourceText(r) { return r.source === 'ai' ? 'AI 답 반영' : r.source === 'template' ? '기본 대책' : '직접 입력'; }
  function reportBook(ctx) {
    var sheets = {}, meta = {};
    var t = ctx.today, rules = cleanRules(ctx.rules);
    var per = ctx.suppliers.map(function (s) {
      var plan = ctx.plans[s] || { rows: [], confirmed: [] };
      return { s: s, plan: plan, an: analyze(ctx.records, { supplier: s, from: plan.from, to: plan.to }, rules), st: planStats(plan.rows || [], t) };
    });
    var rows = []; per.forEach(function (x) { rows = rows.concat(x.an.rows); });
    rows.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.supplier < b.supplier ? -1 : a.supplier > b.supplier ? 1 : 0); });
    var total = analyze(rows, {}, rules);
    // 이슈는 협력사마다 그 협력사 기간으로 계산한 것을 모읍니다(합친 표로 다시 계산하면 기간이 달라짐)
    var issues = []; per.forEach(function (x) { issues = issues.concat(x.an.issues); });

    // 1. 요약
    var sum = [[REPORT_TITLE], ['작성일', t], []];
    var sumHead = sum.length;
    sum.push(['협력사', '분석 기간', '불량 건수', '개선 필요 항목', '개선활동', '완료', '미완료', '보류', '지연', '완료율(%)', '계획 상태']);
    per.forEach(function (x) {
      sum.push([x.s, x.an.from ? x.an.from + ' ~ ' + x.an.to : '', x.an.total, x.an.issues.length, x.st.total, x.st.done, x.st.open, x.st.hold, x.st.overdue, x.st.rate, (x.plan.rows || []).length ? planStateText(x.plan) : 'Master Plan 없음']);
    });
    var sumMeta = { title: [0], head: [sumHead], freeze: { rows: sumHead + 1 } };
    var withAi = per.filter(function (x) { return x.plan.aiSummary; });
    if (withAi.length) {
      sum.push([]); sumMeta.section = [sum.length]; sum.push(['AI 요약 (검토 필요)']);
      sumMeta.head.push(sum.length); sum.push(['협력사', '요약']);
      withAi.forEach(function (x) { sum.push([x.s, x.plan.aiSummary]); });
    }
    sheets['요약'] = sum; meta['요약'] = sumMeta;

    // 2. 불량 목록 — 불량 등록 화면의 칸 순서
    var list = [['No', '발생일자', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치', '분류 근거']];
    rows.forEach(function (r, i) {
      list.push([i + 1, r.date, r.supplier, r.partNo, r.partName, r.type, r.problem, typeof r.qty === 'number' ? r.qty : '', r.action || '', r.manual ? '직접 지정' : r.column ? '불량유형 칸' : r.hits.length ? '키워드: ' + r.hits.join(', ') : '']);
    });
    sheets['불량 목록'] = list;
    meta['불량 목록'] = { head: [0], freeze: { rows: 1 }, filter: 'A1:J' + list.length, widths: [6, 11, 12, 12, 14, 11, 40, 7, 30, 18] };

    // 3. 유형별·협력사별 집계
    var ag = [], am = { section: [], head: [] };
    function sec(title) { if (ag.length) ag.push([]); am.section.push(ag.length); ag.push([title]); }
    function head(cols) { am.head.push(ag.length); ag.push(cols); }
    sec('불량유형별 건수 (파레토)');
    head(['불량유형', '건수', '비율(%)', '누적(%)']);
    total.byType.forEach(function (x) { ag.push([x.type, x.count, x.pct, x.cum]); });
    ag.push(['합계', total.total, total.total ? 100 : 0, '']);
    sec('협력사별 건수');
    head(['협력사', '건수', '개선 필요 항목']);
    total.suppliers.forEach(function (x) { ag.push([x.name, x.count, issues.filter(function (i) { return i.supplier === x.name; }).length]); });
    sec('협력사 × 불량유형');
    var tnames = ctx.types && ctx.types.length ? typeNames(ctx.types) : total.byType.map(function (x) { return x.type; }).filter(function (x) { return x !== UNCLASSIFIED; });
    if (total.unclassified) tnames = tnames.concat([UNCLASSIFIED]);
    head(['협력사'].concat(tnames, ['합계']));
    var cm = {}; rows.forEach(function (r) { var k = r.supplier + '\u0000' + r.type; cm[k] = (cm[k] || 0) + 1; });
    total.suppliers.forEach(function (x) { ag.push([x.name].concat(tnames.map(function (n) { return cm[x.name + '\u0000' + n] || 0; }), [x.count])); });
    ag.push(['합계'].concat(tnames.map(function (n) { var c = 0; total.suppliers.forEach(function (x) { c += cm[x.name + '\u0000' + n] || 0; }); return c; }), [total.total]));
    sec('개선 필요 항목');
    head(['협력사', '불량유형', '건수', '근거']);
    if (!issues.length) ag.push(['(규칙에 걸린 항목 없음)']);
    issues.forEach(function (i) { ag.push([i.supplier, i.type, i.count, i.reasons.join(' / ')]); });
    am.widths = [16, 12, 10, 12];
    sheets['유형별·협력사별 집계'] = ag; meta['유형별·협력사별 집계'] = am;

    // 4. Master Plan — 오른쪽 입력 칸의 항목 순서, 상태 칸은 화면 색
    var mp = [['No', '협력사', '불량유형', '개선대책', '담당부서', '담당자', '계획 시작', '계획 완료', '실적 시작', '실적 완료', '상태', '지연', '근거·비고', '출처']];
    var mpCells = [], all = [];
    per.forEach(function (x) { (x.plan.rows || []).forEach(function (r) { all.push(r); }); });
    all.forEach(function (r, n) {
      mpCells.push({ r: mp.length, c: 10, style: r.status in STATUS ? r.status : 'open' });
      mp.push([n + 1, r.supplier, r.type, r.action, r.dept, r.owner, r.planStart, r.planEnd, r.actStart, r.actEnd, STATUS[r.status] || STATUS.open, isOverdue(r, t) ? '지연' : '', r.note, sourceText(r)]);
    });
    if (!all.length) mp.push(['', '(아직 Master Plan 이 없습니다)']);
    sheets['Master Plan'] = mp;
    meta['Master Plan'] = { head: [0], cells: mpCells, freeze: { rows: 1 }, filter: 'A1:N' + mp.length, widths: [5, 12, 11, 44, 11, 10, 11, 11, 11, 11, 8, 6, 30, 10] };

    // 5. 간트 일정 — 위 두 줄 = 월 / 주차, 칸 색: 계획 = 연한 파랑 ○, 실적 = 상태 색 ●
    var weeks = ganttColumns(all, 'week', { minCols: 1, today: all.length ? '' : t });
    var FIX = ['협력사', '불량유형', '개선대책', '담당부서', '상태'], nf = FIX.length;
    var g0 = FIX.slice(), g1 = FIX.map(function () { return ''; }), merges = [], gc = [], todayCol = -1;
    FIX.forEach(function (f, c) { merges.push({ s: { r: 0, c: c }, e: { r: 1, c: c } }); });
    var col = nf;
    groupSpans(weeks).forEach(function (g) { g0.push(g.group); for (var k = 1; k < g.n; k++) g0.push(''); if (g.n > 1) merges.push({ s: { r: 0, c: col }, e: { r: 0, c: col + g.n - 1 } }); col += g.n; });
    weeks.forEach(function (w, i) { g1.push(w.label); if (w.start <= t && t <= w.end) todayCol = nf + i; });
    var gantt = [g0, g1];
    all.forEach(function (r) {
      var sp = span(r.planStart, r.planEnd, weeks), sa = r.actStart ? span(r.actStart, r.actEnd || (r.status === 'done' ? r.actStart : (t > r.actStart ? t : r.actStart)), weeks) : null;
      var line = [r.supplier, r.type, r.action, r.dept, STATUS[r.status] || STATUS.open], rr = gantt.length;
      gc.push({ r: rr, c: 4, style: r.status in STATUS ? r.status : 'open' });
      weeks.forEach(function (w, i) {
        var p = sp && i >= sp[0] && i <= sp[1], a = sa && i >= sa[0] && i <= sa[1];
        line.push(a ? '●' : p ? '○' : '');
        if (a) gc.push({ r: rr, c: nf + i, style: r.status in STATUS ? r.status : 'open' });
        else if (p) gc.push({ r: rr, c: nf + i, style: 'plan' });
      });
      gantt.push(line);
    });
    if (todayCol >= 0) gc.push({ r: 1, c: todayCol, style: 'today' });
    gantt.push([]);
    gantt.push(['표시: ○ 연한 파랑 = 계획 기간 · ● = 실적 기간(녹색 완료 · 주황 미완료 · 보라 보류) · 노란 머리 칸 = 이번 주. 주차는 목요일이 든 달 기준입니다.']);
    sheets['간트 일정'] = gantt;
    meta['간트 일정'] = { head: [0, 1], cells: gc, merges: merges, freeze: { rows: 2, cols: nf }, widths: [12, 11, 40, 11, 8].concat(weeks.map(function () { return 5; })) };

    // 6. 확정 이력
    var hist = [['협력사', '확정 버전', '확정 일시', '개선활동 수']];
    per.forEach(function (x) { (x.plan.confirmed || []).forEach(function (c) { hist.push([x.s, 'v' + c.version, c.at.replace('T', ' ').slice(0, 16), c.rows.length]); }); });
    sheets['확정 이력'] = hist; meta['확정 이력'] = { head: [0] };
    return { sheets: sheets, meta: meta, weeks: weeks };
  }
  function reportSheets(ctx) { return reportBook(ctx).sheets; }
  /* 엑셀 백업 = 불량 목록을 붙여넣기 양식 순서 그대로(다시 「엑셀 복원」·붙여넣기로 읽힘) */
  function backupSheet(records) {
    var aoa = [PASTE_ORDER.map(fieldLabel)];
    records.forEach(function (r) { aoa.push([r.date, r.supplier, r.partNo, r.partName, r.typeManual || r.typeIn || '', r.problem, typeof r.qty === 'number' ? r.qty : '', r.action || '']); });
    return aoa;
  }

  /* ── 11. 저장 형식 ─────────────────────────────────────────── */
  var SCHEMA_VERSION = 2;
  /* 1단계(v1) 기본 사전 9종 → 수강생 8종. 사전 이름을 손대지 않은 채 저장된 경우에만 바꿔 줍니다 */
  var V1_TYPES = { '이종 불량': '이종/누락', '용접 균열': '용접불량', '용접 불량': '용접불량', '가공 치수불량': '위치/치수', '도장 불량': '기타',
    '외관 불량': '기타', '조립·체결 불량': '미조임', '누유·누수': '누유/누수', '녹·부식': '기타' };
  function isV1Defaults(types) {
    var names = (types || []).map(function (t) { return t && str(t.name); });
    return names.length === 9 && names.every(function (n) { return V1_TYPES[n]; });
  }
  function emptyDb() {
    return { schemaVersion: SCHEMA_VERSION, records: [], seq: 0, types: defaultTypes(), depts: DEFAULT_DEPTS.slice(), rules: cleanRules({}),
      plans: {}, aliases: { suppliers: {}, parts: {} }, mask: true, mapping: null, files: [] };
  }
  function cleanType(t) {
    if (!t || typeof t !== 'object' || !str(t.name) || str(t.name) === UNCLASSIFIED) return null;
    return { name: str(t.name), dept: str(t.dept) || '협력사', keywords: Array.isArray(t.keywords) ? uniq(t.keywords.map(str).filter(Boolean)) : [],
      actions: Array.isArray(t.actions) ? t.actions.map(function (a) { return a && str(a.text) ? { text: str(a.text), dept: str(a.dept) || '부품품질팀', start: Math.max(0, parseInt(a.start, 10) || 0), weeks: Math.max(1, parseInt(a.weeks, 10) || 1) } : null; }).filter(Boolean) : [] };
  }
  function restoreDb(o) {
    var db = emptyDb();
    if (!o || typeof o !== 'object') return db;
    if (Array.isArray(o.records)) db.records = o.records.filter(function (r) { return r && typeof r === 'object'; }).map(function (r) {
      var c = { id: str(r.id), date: parseDate(r.date), partName: str(r.partName), partNo: str(r.partNo), problem: str(r.problem), supplier: str(r.supplier) };
      if (typeof r.qty === 'number') c.qty = r.qty;
      if (str(r.typeManual)) c.typeManual = str(r.typeManual);
      if (str(r.typeIn)) c.typeIn = str(r.typeIn);
      if (str(r.action)) c.action = str(r.action);
      return c;
    });
    if (typeof o.seq === 'number') db.seq = o.seq;
    db.records.forEach(function (r) { if (!r.id) r.id = 'd' + (++db.seq); });
    var v1 = Array.isArray(o.types) && !(o.schemaVersion >= 2) && isV1Defaults(o.types);
    if (Array.isArray(o.types) && !v1) { var ts = o.types.map(cleanType).filter(Boolean); if (ts.length) db.types = ts; }
    if (v1) db.records.forEach(function (r) { if (r.typeManual && V1_TYPES[r.typeManual]) r.typeManual = V1_TYPES[r.typeManual]; });
    if (Array.isArray(o.depts)) { var ds = uniq(o.depts.map(str).filter(Boolean)); if (ds.length) db.depts = ds; }
    db.rules = cleanRules(o.rules);
    if (o.plans && typeof o.plans === 'object') Object.keys(o.plans).forEach(function (s) {
      var p = o.plans[s]; if (!p || typeof p !== 'object') return;
      db.plans[s] = { baseDate: parseDate(p.baseDate), from: parseDate(p.from), to: parseDate(p.to), rows: Array.isArray(p.rows) ? p.rows.map(cleanPlanRow) : [],
        confirmed: Array.isArray(p.confirmed) ? p.confirmed.filter(function (c) { return c && Array.isArray(c.rows); }).map(function (c, i) { return { version: i + 1, at: str(c.at), rows: c.rows.map(cleanPlanRow) }; }) : [],
        aiSummary: str(p.aiSummary), aiChecks: Array.isArray(p.aiChecks) ? p.aiChecks.map(str).filter(Boolean) : [], aiAnswer: p.aiAnswer == null ? '' : String(p.aiAnswer) };
    });
    if (v1) Object.keys(db.plans).forEach(function (s) {
      var p = db.plans[s];
      [p.rows].concat(p.confirmed.map(function (c) { return c.rows; })).forEach(function (rs) { rs.forEach(function (r) { if (V1_TYPES[r.type]) r.type = V1_TYPES[r.type]; }); });
    });
    if (o.aliases && typeof o.aliases === 'object') db.aliases = { suppliers: assignAliases(o.aliases.suppliers, [], ALIAS.suppliers), parts: assignAliases(o.aliases.parts, [], ALIAS.parts) };
    db.mask = o.mask !== false;
    if (o.mapping && typeof o.mapping === 'object') db.mapping = o.mapping;
    if (Array.isArray(o.files)) db.files = o.files.slice(-50);
    if (o._sample) db._sample = true;
    return db;
  }
  function suppliersOf(records) { return uniq(records.map(function (r) { return r.supplier; }).filter(Boolean)).sort(); }

  return {
    // 날짜
    parseDate: parseDate, toDateStr: toDateStr, addDays: addDays, daysBetween: daysBetween, mondayOf: mondayOf, monthOf: monthOf, monthAdd: monthAdd, weekInfo: weekInfo,
    // 읽기
    FIELDS: FIELDS, PASTE_ORDER: PASTE_ORDER, fieldLabel: fieldLabel, detectColumns: detectColumns, parseSheet: parseSheet, pickSheet: pickSheet, recordKey: recordKey, mergeRecords: mergeRecords, normHead: normHead,
    splitTsv: splitTsv, parsePaste: parsePaste, checkRecordInput: checkRecordInput, recordFromInput: recordFromInput,
    // 분류
    UNCLASSIFIED: UNCLASSIFIED, DEFAULT_DEPTS: DEFAULT_DEPTS, defaultTypes: defaultTypes, classify: classify, typeOf: typeOf, enrich: enrich, typeNames: typeNames, findTypeName: findTypeName,
    parseActionLines: parseActionLines, formatActionLines: formatActionLines, parseKeywords: parseKeywords,
    // 분석
    DEFAULT_RULES: DEFAULT_RULES, cleanRules: cleanRules, analyze: analyze, findIssues: findIssues, dateBounds: dateBounds, suppliersOf: suppliersOf,
    // 계획
    STATUS: STATUS, buildDraft: buildDraft, makeRow: makeRow, cleanPlanRow: cleanPlanRow, checkPlanRow: checkPlanRow, planStats: planStats, isOverdue: isOverdue,
    ganttColumns: ganttColumns, span: span, groupSpans: groupSpans,
    confirmPlan: confirmPlan, lastConfirmed: lastConfirmed, planChangedSince: planChangedSince, planStateText: planStateText,
    // AI
    ALIAS: ALIAS, assignAliases: assignAliases, updateAliases: updateAliases, maskText: maskText, unmaskText: unmaskText, findLeaks: findLeaks,
    buildAiPromptInfo: buildAiPromptInfo, parseAiPlan: parseAiPlan, matchType: matchType,
    // 보고서·저장
    reportSheets: reportSheets, reportBook: reportBook, backupSheet: backupSheet, REPORT_TITLE: REPORT_TITLE, emptyDb: emptyDb, restoreDb: restoreDb, SCHEMA_VERSION: SCHEMA_VERSION
  };
});
