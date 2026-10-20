// 가상 샘플 파일 만들기: node scripts/make-samples.js
// js/sample-data.js 의 예시(협력사 A~F, 전부 지어낸 값)를 엑셀·CSV 로 씁니다.
// 날짜 칸은 엑셀 날짜로 넣고, 방금 쓴 파일을 앱과 같은 방식으로 다시 읽어 원본과 같은지 확인합니다.
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');

const out = path.join(__dirname, '..', 'samples');
fs.mkdirSync(out, { recursive: true });
const NAME = '예시_협력사_공정불량_RawData';
const aoa = Sample.aoa().map((row, i) => i < 3 ? row : row.map((v, c) => c === 1 ? new Date(v + 'T00:00:00') : v));

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'yyyy-mm-dd' });
ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
ws['!cols'] = [{ wch: 6 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 34 }, { wch: 6 }];
XLSX.utils.book_append_sheet(wb, ws, '공정불량 RAW');
// 안내 시트 — 앱은 머리행이 맞는 「공정불량 RAW」 시트를 고릅니다
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['이 파일은 도구 시험용 가상 데이터입니다.'], ['협력사 A~F · 품번 TP-… 는 모두 지어낸 이름입니다.']]), '안내');
const xlsxPath = path.join(out, NAME + '.xlsx');
fs.writeFileSync(xlsxPath, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellDates: true }));

// CSV — 날짜를 「2026.03.05」 식으로 적어 다른 표기도 읽히는지 봅니다
const csv = Sample.aoa().slice(2).map((row, i) => row.map((v, c) => {
  let s = v == null ? '' : String(v);
  if (i > 0 && c === 1) s = s.replace(/-/g, '.');
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}).join(',')).join('\r\n');
const csvPath = path.join(out, NAME + '.csv');
fs.writeFileSync(csvPath, '﻿' + csv);

// 다시 읽어 확인
const want = L.mergeRecords([], Sample.records(), 0).records.map(L.recordKey);
for (const [p, opt] of [[xlsxPath, { type: 'buffer' }], [csvPath, { type: 'string' }]]) {
  const buf = opt.type === 'string' ? fs.readFileSync(p, 'utf8').replace(/^﻿/, '') : fs.readFileSync(p);
  const book = XLSX.read(buf, opt);
  const sheets = book.SheetNames.map(n => ({ name: n, aoa: XLSX.utils.sheet_to_json(book.Sheets[n], { header: 1, raw: true, defval: null }) }));
  const pick = L.pickSheet(sheets);
  const r = L.parseSheet(pick.aoa);
  const got = L.mergeRecords([], r.records, 0).records.map(L.recordKey);
  if (r.det.missing.length || r.badDate || JSON.stringify(got) !== JSON.stringify(want)) {
    throw new Error(path.basename(p) + ' 다시 읽기 불일치: missing=' + r.det.missing + ' badDate=' + r.badDate + ' 건수 ' + got.length + '/' + want.length);
  }
  console.log(path.basename(p), '시트', pick.name, '·', got.length + '건 — 원본과 같음');
}
