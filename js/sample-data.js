/* 예시 데이터 — 전부 지어낸 가상 값입니다(협력사 A~F, 품번 TP-…). 실제 회사·협력사 자료가 아닙니다.
   고정 씨앗 난수라 늘 같은 데이터가 나옵니다. 분석 규칙이 제대로 잡는지 보려고 일부러 심어 둔 것:
   - 협력사 C : 용접 균열이 8~10월에 늘어남(증가 추세), 같은 품번 TP-3301 반복
   - 협력사 A : 이종 불량 반복, 같은 품번 TP-1102 반복
   - 협력사 D : 가공 치수불량 반복, 같은 품번 TP-4405 반복
   - 협력사 F : 건수가 적어 규칙에 걸리지 않음(초안은 건수 상위 유형으로)
   - 어느 유형에도 맞지 않는 문제점 몇 건(미분류) */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SQSample = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  function rng(seed) {
    return function () {
      seed |= 0; seed = seed + 0x6D2B79F5 | 0;
      var t = Math.imul(seed ^ seed >>> 15, 1 | seed);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }
  var TEXT = {
    '이종 불량': ['이종품 혼입 (타 품번 2개 섞임)', '오품 납품 — 좌우 품번 바뀜', '라벨 오류로 품번 상이', '이종 볼트 혼용', '사양 상이 부품 혼입'],
    '용접 균열': ['용접부 균열 발생', '용접 비드 끝단 크랙', '브래킷 용접부 갈라짐', '용접 후 균열(열영향부)'],
    '용접 불량': ['용접 기공 다수', '언더컷 발생', '스패터 과다 부착', '미용접 구간 있음', '각장 부족'],
    '가공 치수불량': ['홀 위치 치수 공차 초과', '가공 치수 규격 미달(외경)', '탭 나사산 불량', '평면도 불량', '가공면 burr 잔존'],
    '도장 불량': ['도막 두께 부족', '도장 흐름 발생', '도장 박리', '미도장 부위 있음', '도장면 기포'],
    '외관 불량': ['표면 찍힘', '스크래치 다수', '이물 부착', '운송 중 변형(휨)'],
    '조립·체결 불량': ['볼트 체결 토크 부족', '와셔 누락', '오조립(방향 반대)', '너트 풀림'],
    '누유·누수': ['유압 파이프 연결부 누유', '기밀 시험 리크 발생', '피팅부 오일 샘'],
    '녹·부식': ['보관 중 녹 발생', '도금면 부식']
  };
  var OTHER = ['검사성적서 미첨부', '납품 수량 부족', '원인 확인 중'];
  var SUP = {
    '협력사 A': { n: 26, parts: [['TP-1101', '브래킷'], ['TP-1102', '브래킷 ASSY'], ['TP-1103', '커버'], ['TP-1104', '핀']], w: { '이종 불량': 5, '외관 불량': 2, '가공 치수불량': 1, '조립·체결 불량': 1 } },
    '협력사 B': { n: 30, parts: [['TP-2201', '탱크'], ['TP-2202', '커버'], ['TP-2203', '도어 판넬']], w: { '도장 불량': 5, '외관 불량': 3, '녹·부식': 1 } },
    '협력사 C': { n: 22, parts: [['TP-3301', '프레임'], ['TP-3302', '브래킷'], ['TP-3303', '서포트']], w: { '용접 불량': 4, '가공 치수불량': 1, '외관 불량': 1 } },
    '협력사 D': { n: 28, parts: [['TP-4401', '샤프트'], ['TP-4402', '부싱'], ['TP-4403', '하우징'], ['TP-4405', '플랜지']], w: { '가공 치수불량': 5, '외관 불량': 1, '녹·부식': 1 } },
    '협력사 E': { n: 24, parts: [['TP-5501', '유압 파이프'], ['TP-5502', '호스 ASSY'], ['TP-5503', '밸브 블록']], w: { '조립·체결 불량': 3, '누유·누수': 3, '외관 불량': 1 } },
    '협력사 F': { n: 3, parts: [['TP-6601', '가이드'], ['TP-6602', '스페이서']], w: { '외관 불량': 1, '가공 치수불량': 1 } }
  };
  var START = '2026-01-05', END = '2026-10-16';
  function addDays(s, n) { var d = new Date(Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) + n * 86400000); return d.toISOString().slice(0, 10); }
  function span() { return Math.round((Date.UTC(2026, 9, 16) - Date.UTC(2026, 0, 5)) / 86400000); }
  function pick(r, arr) { return arr[Math.floor(r() * arr.length)]; }
  function wpick(r, w) {
    var ks = Object.keys(w), tot = ks.reduce(function (s, k) { return s + w[k]; }, 0), x = r() * tot;
    for (var i = 0; i < ks.length; i++) { x -= w[ks[i]]; if (x < 0) return ks[i]; }
    return ks[ks.length - 1];
  }
  /* 발생일자 순 레코드 [{date, supplier, partNo, partName, problem, qty}] */
  function records() {
    var r = rng(20261020), out = [], days = span();
    Object.keys(SUP).forEach(function (s) {
      var p = SUP[s];
      for (var i = 0; i < p.n; i++) {
        var t = wpick(r, p.w), part = pick(r, p.parts);
        out.push({ date: addDays(START, Math.floor(r() * days)), supplier: s, partNo: part[0], partName: part[1], problem: pick(r, TEXT[t]), qty: 1 + Math.floor(r() * 6) });
      }
    });
    // 심어 둔 신호
    function put(date, s, pn, name, problem, qty) { out.push({ date: date, supplier: s, partNo: pn, partName: name, problem: problem, qty: qty }); }
    put('2026-03-11', '협력사 C', 'TP-3302', '브래킷', '용접부 균열 발생', 2);
    ['2026-08-04', '2026-08-21', '2026-09-02', '2026-09-17', '2026-09-29', '2026-10-08', '2026-10-14'].forEach(function (d, i) {
      put(d, '협력사 C', i % 3 === 2 ? 'TP-3303' : 'TP-3301', i % 3 === 2 ? '서포트' : '프레임', TEXT['용접 균열'][i % 4], 1 + (i % 3));
    });
    ['2026-02-17', '2026-05-12', '2026-07-08', '2026-09-23'].forEach(function (d, i) { put(d, '협력사 A', 'TP-1102', '브래킷 ASSY', TEXT['이종 불량'][i % 3], 4 + i); });
    ['2026-04-06', '2026-06-15', '2026-08-26'].forEach(function (d, i) { put(d, '협력사 D', 'TP-4405', '플랜지', TEXT['가공 치수불량'][i], 3); });
    put('2026-05-20', '협력사 B', 'TP-2202', '커버', OTHER[0], 1);
    put('2026-07-14', '협력사 E', 'TP-5502', '호스 ASSY', OTHER[1], 10);
    put('2026-09-09', '협력사 F', 'TP-6601', '가이드', OTHER[2], 1);
    out.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.supplier < b.supplier ? -1 : 1; });
    return out;
  }
  /* 엑셀 모양 — 제목 줄과 빈 줄이 있고, 머리행 이름을 일부러 표준과 다르게(업체명·품 번·불량 내용) 적었습니다 */
  var HEAD = ['No', '발생일', '업체명', '품 번', '품명', '불량 내용', '수량'];
  function aoa() {
    var rows = [['■ 협력사 공정불량 현황 (가상 예시 — 실제 회사 자료 아님)'], [], HEAD];
    records().forEach(function (x, i) { rows.push([i + 1, x.date, x.supplier, x.partNo, x.partName, x.problem, x.qty]); });
    return rows;
  }
  return { records: records, aoa: aoa, HEAD: HEAD, START: START, END: END };
});
