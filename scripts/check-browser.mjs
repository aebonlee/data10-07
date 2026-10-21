// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어 불량 등록(직접 입력·엑셀 붙여넣기·수정·엑셀 백업) → 예시 데이터 → 불량 분석
// → Master Plan 초안 → 오른쪽 칸 실적 입력 → 계획 확정 → AI 답 반영 → 추진현황 → 보고서 내려받기(서식 확인) → 설정 순서로 누르고,
// JS 오류 · 외부 주소 요청(사내망 전용이라 0건이어야 함) · 폭별 가로 넘침을 잽니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const url = 'file://' + path.join(root, 'index.html');

const errors = [];
const problems = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
page.on('dialog', d => d.accept());
const external = [];
page.on('request', r => { const u = r.url(); if (!/^(file|data|blob):/.test(u)) external.push(u); });

function ok(cond, msg) { if (!cond) problems.push(msg); else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(150); }
async function clickText(sel, text) { await page.locator(sel, { hasText: text }).first().click(); await page.waitForTimeout(150); }

await page.goto(url);
// 점검기 자체 확인용: BREAK=1 이면 일부러 넓은 상자를 넣어 넘침을 잡는지 봅니다
if (process.env.BREAK) await page.addStyleTag({ content: '.page-head { min-width: 900px; }' });
// 점검기 자체 확인용: NET=1 이면 일부러 외부 주소를 불러 요청 점검이 잡는지 봅니다
if (process.env.NET) await page.evaluate(() => { const i = new Image(); i.src = 'https://example.com/x.png'; });

// ── 불량 등록 (파일 없이) ──
await go('#/entry');
const f = '#entryForm';
await page.fill(f + ' input[type=date]', '2026-10-05');
await page.fill(f + ' input[list="dl-sup"]', '협력사 가');
await page.fill(f + ' input[placeholder="예: TP-0001"]', 'TP-9001');
await page.fill(f + ' textarea >> nth=0', '브래킷 용접부 균열 발생');
ok((await page.textContent('#entryCard')).includes('자동 분류: 용접불량'), '내용 입력 중 자동 분류 안내(용접불량)');
await page.fill(f + ' input[type=number]', '2');
await page.fill(f + ' textarea >> nth=1', '전수 선별');
await clickText(f + ' button', '등록');
ok(await page.inputValue(f + ' input[list="dl-sup"]') === '협력사 가', '등록 후 협력사 남겨 둠');
await page.fill(f + ' textarea >> nth=0', '호스 체결부 조임 불량');
await page.selectOption(f + ' select', '미조임');
await clickText(f + ' button', '등록');
await page.fill(f + ' textarea >> nth=0', '');
await clickText(f + ' button', '등록');
ok((await page.textContent('#entryCard')).includes('내용(문제점)을 적어 주세요'), '필수 칸 빠지면 안내');
// 엑셀 붙여넣기 — 머리행 없음(순서대로) 2줄 + 유형 칸 「위치·치수」
await page.fill('textarea.paste-box', '2026-10-06\t협력사 나\tTP-9002\t커버\t위치·치수\t홀 위치 틀림\t1\t재작업\n2026.10.07\t협력사 나\tTP-9003\t\t\t오일 샘 발생\t\t');
await clickText('button', '표 읽기');
ok((await page.textContent('.ai-preview')).includes('읽은 결과 — 2건'), '붙여넣기 2건 읽음');
await clickText('button', '2건 넣기');
const listText = await page.textContent('main');
ok(listText.includes('3. 등록된 불량 4건'), '등록된 불량 4건');
ok(listText.includes('위치/치수') && listText.includes('누유/누수') && listText.includes('미조임'), '유형: 위치/치수(유형 칸) · 누유/누수(키워드) · 미조임(직접 고름)');
// 수정
await page.locator('main table.list tbody tr', { hasText: '오일 샘' }).locator('button', { hasText: '수정' }).click();
await page.waitForTimeout(150);
await page.fill(f + ' textarea >> nth=1', '협력사 반품');
await clickText(f + ' button', '수정 저장');
ok((await page.locator('main table.list tbody tr', { hasText: '오일 샘' }).textContent()).includes('협력사 반품'), '수정 저장(조치)');
// 엑셀 백업 → 다시 읽기
{
  const [d] = await Promise.all([page.waitForEvent('download'), clickText('button', '엑셀 백업 받기')]);
  const fp = path.join(root, 'samples', '.check-backup.xlsx'); await d.saveAs(fp);
  const b = XLSX.read(fs.readFileSync(fp), { type: 'buffer' }); fs.unlinkSync(fp);
  const rows = XLSX.utils.sheet_to_json(b.Sheets['불량 목록'], { header: 1, raw: false });
  ok(rows.length === 5 && rows[0].join(',') === '발생일자,협력사,품번,품명,불량유형,내용,수량,조치', '엑셀 백업 4건 + 머리행');
}
// 예시 데이터로 바꾸기(확인 창)
await go('#/upload');
await clickText('button', '예시 데이터로 시작');
await page.locator('#dialog button', { hasText: '바꾸기' }).click();
await page.waitForTimeout(200);
ok(await page.locator('#sampleBanner').isVisible(), '예시 데이터 띠 보임');
ok((await page.textContent('main')).includes('개선 필요 항목'), '불량 분석 화면');

// 분류 수동 지정
const sel = page.locator('select.type-sel').first();
await sel.selectOption('기타');
await page.waitForTimeout(100);
ok(await page.locator('select.type-sel.manual').count() >= 1, '불량유형 직접 고르기 저장');

// Master Plan
await go('#/plan/' + encodeURIComponent('협력사 C'));
await clickText('button', '분석 및 초안 생성');
const nRows = await page.locator('table.gantt tbody tr.clickable').count();
ok(nRows > 5, 'Master Plan 초안 행 ' + nRows + '개');
ok((await page.textContent('table.gantt thead')).includes('2026년'), '간트 머리에 연·월');
ok(await page.locator('.dkpis.six .dk').count() === 6, 'KPI 6칸');
ok(!(await page.locator('#aiCard').getAttribute('open') !== null), 'AI 칸은 기본으로 접힘(선택 기능)');

// 오른쪽 칸 — 실적·완료
await page.locator('table.gantt tbody tr.clickable').first().click();
ok(await page.locator('#drawer').isVisible(), '행 누르면 오른쪽 칸');
await page.locator('#drawer input[placeholder^="예: 홍길동"]').fill('김담당 책임');
await page.locator('#drawer .st-btn.st-done').click();
await clickText('#drawer button', '저장');
ok(await page.locator('#drawer').isHidden(), '저장 후 닫힘');
ok(await page.locator('table.gantt .pill.st-done').count() === 1, '완료 상태 1건');
ok(await page.locator('table.gantt .g-act.st-done').count() === 1, '실적 막대(녹색)');

// 확정
await clickText('button', '계획 확정');
await page.locator('#dialog button', { hasText: '확정' }).last().click();
await page.waitForTimeout(150);
ok((await page.textContent('.gantt-card')).includes('확정 v1'), '계획 확정 v1');

// AI 답 반영(선택 기능 — 칸을 펼쳐서)
await page.locator('#aiCard > summary').click();
const prompt = await page.locator('textarea.prompt-box[readonly]').inputValue();
ok(!/협력사 [A-F]|TP-\d{4}/.test(prompt), '프롬프트에 협력사·품번 없음(별칭)');
const alias = (/협력사: (업체\d+)/.exec(prompt) || [])[1];
await page.locator('textarea.prompt-box.answer').fill(['요약: ' + alias + ' 용접 균열이 늘었어.', '| 불량유형 | 개선대책 | 담당부서 | 시작 | 기간 |', '|---|---|---|---|---|', '| 용접 불량 | 예열 온도 기록지 도입 | 협력사 | 0 | 2 |', '| 위치·치수 | 지그 점검표 개정 | 생산기술팀 | 1 | 3 |', '확인 필요: 예열 설비 유무'].join('\n'));
await clickText('button', '답 읽기');
ok((await page.textContent('.ai-preview')).includes('개선대책 2건'), 'AI 답 2건 읽음');
await clickText('button', '계획에 반영');
ok((await page.textContent('.ai-sum')).includes('협력사 C 용접 균열이 늘었어'), 'AI 요약 별칭 되돌림');
ok((await page.textContent('table.gantt')).includes('위치/치수'), 'AI 답의 「위치·치수」 → 위치/치수');
ok((await page.textContent('.gantt-card')).includes('이후 계획 수정 있음'), '확정 뒤 수정 표시');

// 개선활동 추가
await clickText('button', '개선활동 추가');
await page.locator('#drawer textarea').first().fill('협력사 품질회의 정례화');
await page.locator('#drawer .st-btn.st-hold').click();
await clickText('#drawer button', '저장');
ok(await page.locator('table.gantt .pill.st-hold').count() === 1, '보류 활동 추가');

// 추진현황·보고서
await go('#/progress');
ok(await page.locator('table.gantt tbody tr.clickable').count() > 0, '추진현황 간트');
await go('#/report');
const [dl] = await Promise.all([page.waitForEvent('download'), clickText('button', 'Excel 보고서')]);
const file = path.join(root, 'samples', '.check-report.xlsx');
await dl.saveAs(file);
const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellNF: true });
if (process.env.KEEP) fs.copyFileSync(file, process.env.KEEP);
fs.unlinkSync(file);
ok(wb.SheetNames.join(',') === '요약,불량 목록,유형별·협력사별 집계,Master Plan,간트 일정,확정 이력', '보고서 시트 6개: ' + wb.SheetNames.join(','));
const mp = XLSX.utils.sheet_to_json(wb.Sheets['Master Plan'], { header: 1 });
ok(mp.some(r => r[3] === '협력사 품질회의 정례화' && r[10] === '보류'), '보고서에 보류 활동');
ok(wb.Sheets['불량 목록'].B2.t === 'n' && wb.Sheets['불량 목록'].B2.z === 'yyyy-mm-dd', '불량 목록 발생일자 = 엑셀 날짜');
const g = XLSX.utils.sheet_to_json(wb.Sheets['간트 일정'], { header: 1, raw: false });
ok(g[0][5] && /^\d+주$/.test(g[1][5]) && g.some(r => r.includes('○')) && g.some(r => r[2] === '협력사 품질회의 정례화' && r[4] === '보류'), '간트 일정 시트 — 월/주차 머리 · ○ 계획 · 보류 활동');

// 설정·대시보드
await go('#/settings');
ok(await page.locator('.dict-item').count() === 8, '불량유형 사전 8종');
ok((await page.textContent('main')).includes('사내망에서 쓰기'), '사내망 안내');
await go('#/home');
ok((await page.textContent('main')).includes('협력사별 추진현황'), '대시보드');

// 폭별 가로 넘침 — 안쪽 가로 스크롤 상자(overflow-x auto) 안은 의도된 것이라 뺍니다
for (const w of [360, 390, 768, 1024, 1280, 1440]) {
  await page.setViewportSize({ width: w, height: 900 });
  for (const r of ['#/home', '#/entry', '#/upload', '#/analysis', '#/plan', '#/progress', '#/report', '#/settings']) {
    await go(r);
    const bad = await page.evaluate(() => {
      const cw = document.documentElement.clientWidth, out = [];
      for (const el of document.querySelectorAll('main *, .topbar *')) {
        const b = el.getBoundingClientRect();
        if (!b.width || !b.height) continue;
        let p = el.parentElement, inner = false;
        while (p) { const ox = getComputedStyle(p).overflowX; if (ox === 'auto' || ox === 'scroll' || ox === 'hidden') { inner = true; break; } p = p.parentElement; }
        if (!inner && b.right > cw + 1) out.push(el.tagName + '.' + el.className + ' ' + Math.round(b.right) + '>' + cw);
      }
      return { out: out.slice(0, 3), page: document.documentElement.scrollWidth > cw + 1 };
    });
    if (bad.page || bad.out.length) problems.push(w + 'px ' + r + ' 넘침: ' + bad.out.join(' | '));
  }
}
console.log('  폭 6가지 × 화면 8개 넘침 점검 끝');
if (external.length) problems.push('외부 주소 요청 ' + external.length + '건: ' + [...new Set(external)].slice(0, 3).join(', '));
else console.log('  ok  외부 주소 요청 0건(사내망 전용)');

await browser.close();
if (errors.length) problems.push(...errors);
if (problems.length) { console.error('\n문제 ' + problems.length + '건\n' + problems.join('\n')); process.exitCode = 1; }
else console.log('\n브라우저 점검 통과');
