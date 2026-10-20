/* Quality Planner — 화면 (대시보드 · RAW DATA 업로드 · 불량 분석 · MASTER PLAN · 추진현황 · 보고서 출력 · 설정)
   계산은 js/logic.js, 저장은 js/store.js, AI 연결은 js/ai-endpoint.js · js/ai-panel.js 에 있습니다. */
(function () {
  'use strict';
  var L = window.SQLogic;
  var S = window.SQStore;
  var Sample = window.SQSample;
  var XLSX = window.XLSX;
  var PAGE = 50;
  var COLOR = { done: '#1f9d55', open: '#e8830c', hold: '#8e44ad' };

  var db = S.loadDb();
  var ui = {
    an: { supplier: '', from: '', to: '', type: '', q: '', onlyUnc: false, page: 0 },
    plan: { supplier: '', scale: 'week', aiText: '', parsed: null, mode: 'replace', showPrompt: false },
    prog: { supplier: '', status: '', overdue: false, scale: 'month' },
    rep: { supplier: '' },
    imp: null,
    impMode: 'add'
  };

  // ── 작은 도구 ─────────────────────────────────────────────
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k === 'value') el.value = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(el, arguments[i]);
    return el;
  }
  function add(el, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(el, x); }); return; }
    el.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
  }
  function fmt(n) { return n == null || n === '' ? '' : Number(n).toLocaleString('ko-KR', { maximumFractionDigits: 1 }); }
  function toast(msg, isError) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast' + (isError ? ' error' : '');
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, 3800);
  }
  function save() {
    if (!S.saveDb(db)) {
      document.getElementById('storeBanner').hidden = false;
      if (S.lastError() === 'quota') toast('브라우저 저장 공간이 찼습니다. 설정에서 JSON 백업을 받아 두세요.', true);
    }
  }
  function today() { return L.toDateStr(new Date()); }
  function nowIso() { var d = new Date(); return today() + 'T' + ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2) + ':' + ('0' + d.getSeconds()).slice(-2); }
  function nextMonday() { var t = today(), m = L.mondayOf(t); return m === t ? t : L.addDays(m, 7); }
  function field(label, input, hint, cls) {
    return h('label', { class: 'field' + (cls ? ' ' + cls : '') }, h('span', null, label), input, hint ? h('small', { class: 'hint' }, hint) : null);
  }
  function select(options, value, attrs) {
    var s = h('select', attrs || {});
    options.forEach(function (o) {
      var v = Array.isArray(o) ? o[0] : o, t = Array.isArray(o) ? o[1] : o;
      s.appendChild(h('option', { value: v, selected: String(v) === String(value) }, t));
    });
    return s;
  }
  function dateInput(value, attrs) { return h('input', Object.assign({ type: 'date', value: value || '' }, attrs || {})); }
  function datalist(id, values) { return h('datalist', { id: id }, values.map(function (v) { return h('option', { value: v }); })); }
  var onDialogClose = null;
  function openDialog(title, content, actions) {
    var d = document.getElementById('dialog');
    document.getElementById('dialogTitle').textContent = title;
    var c = document.getElementById('dialogContent'); c.textContent = ''; add(c, content);
    var a = document.getElementById('dialogActions'); a.textContent = ''; add(a, actions);
    if (!d.open) d.showModal();
  }
  function closeDialog() { var d = document.getElementById('dialog'); if (d.open) d.close(); }
  document.getElementById('dialog').addEventListener('close', function () { var f = onDialogClose; onDialogClose = null; if (f) f(); });
  function confirmBox(title, text, okLabel, onOk, danger) {
    openDialog(title, h('p', { class: 'pre' }, text), [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '그만두기'),
      h('button', { type: 'button', class: 'btn ' + (danger ? 'btn-danger' : 'btn-primary'), onclick: function () { closeDialog(); onOk(); } }, okLabel)
    ]);
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function writeXlsx(name, sheets, widths) {
    var wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(function (n) {
      var ws = XLSX.utils.aoa_to_sheet(sheets[n]);
      var w = (widths && widths[n]) || [];
      var cols = 0; sheets[n].forEach(function (r) { cols = Math.max(cols, r.length); });
      ws['!cols'] = Array.from({ length: cols }, function (_, i) { return { wch: w[i] || 14 }; });
      XLSX.utils.book_append_sheet(wb, ws, n.replace(/[\\\/?*\[\]:]/g, '-').slice(0, 31));
    });
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    download(name, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  function copyText(s) {
    function fallback() {
      var ta = h('textarea', { style: 'position:fixed;left:-9999px;top:0' });
      ta.value = s; document.body.appendChild(ta); ta.select();
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      ta.remove();
      toast(ok ? '복사했습니다' : '복사하지 못했습니다. 글상자를 눌러 직접 복사하세요', !ok);
    }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(s).then(function () { toast('복사했습니다'); }, fallback);
    else fallback();
  }
  function tag() { return db._sample ? '_예시데이터' : ''; }
  function all() { return L.enrich(db.records, db.types); }
  function suppliers() { return L.suppliersOf(db.records); }
  function statusPill(st, overdue) {
    return h('span', { class: 'pill st-' + st }, L.STATUS[st] || '미완료', overdue ? h('span', { class: 'late' }, ' · 지연') : null);
  }
  function emptyNotice(main, title) {
    main.appendChild(pageHead(title));
    main.appendChild(h('section', { class: 'card' },
      h('p', null, '아직 공정불량 Raw data 가 없습니다. 「RAW DATA 업로드」에서 Excel 을 올리거나, 예시 데이터로 흐름을 먼저 볼 수 있습니다.'),
      h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/upload' }, 'RAW DATA 업로드로'), sampleButton(false))));
  }
  function pageHead(title, sub, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'ph-text' }, h('h1', null, title), sub ? h('p', { class: 'note' }, sub) : null), right || null);
  }

  // ── 아이콘·차트 부품 ──────────────────────────────────────
  var SVGNS = 'http://www.w3.org/2000/svg';
  function sv(tagName, attrs) {
    var el = document.createElementNS(SVGNS, tagName);
    Object.keys(attrs || {}).forEach(function (k) { el.setAttribute(k, attrs[k]); });
    for (var i = 2; i < arguments.length; i++) { var c = arguments[i]; if (c == null) continue; el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); }
    return el;
  }
  var ICON = {
    doc: 'M6 3h9l4 4v14H6zM9 12h7M9 16h5M14 3v5h5',
    factory: 'M3 20V10l5 3V10l5 3V6h3l1 4h3v10zM7 17h2M12 17h2M17 17h1',
    alert: 'M12 4 2.8 19.5h18.4zM12 10v4.5M12 17.2v.3',
    check: 'M4 4h16v16H4zM8 12.5l3 3 5-6',
    chev: 'm9 6 6 6-6 6',
    tick: 'm5 12.5 4.5 4.5L19 7.5',
    chip: 'M7 7h10v10H7zM10 10h4v4h-4zM9.5 3v4M14.5 3v4M9.5 17v4M14.5 17v4M3 9.5h4M3 14.5h4M17 9.5h4M17 14.5h4',
    layers: 'M12 3 3 8l9 5 9-5zM3 13l9 5 9-5M3 17.5l9 5 9-5',
    list: 'M4 6h16M4 12h16M4 18h10',
    cal: 'M4 6h16v14H4zM4 10h16M8 3v5M16 3v5',
    clock: 'M12 3.5a8.5 8.5 0 1 0 0 17a8.5 8.5 0 1 0 0-17M12 7.5V12l3 2',
    pause: 'M12 3.5a8.5 8.5 0 1 0 0 17a8.5 8.5 0 1 0 0-17M10 9v6M14 9v6',
    plus: 'M12 5v14M5 12h14',
    down: 'M12 4v11M7.5 10.5 12 15l4.5-4.5M4 19h16',
    redo: 'M4 12a8 8 0 0 1 14-5.3L20 9M20 4v5h-5M20 12a8 8 0 0 1-14 5.3L4 15M4 20v-5h5'
  };
  function icon(name, cls) { return sv('svg', { viewBox: '0 0 24 24', class: 'ico' + (cls ? ' ' + cls : ''), 'aria-hidden': 'true' }, sv('path', { d: ICON[name] })); }
  function kpi(title, iconName, value, unit, note, tone) {
    return h('div', { class: 'dk' + (tone ? ' tone-' + tone : '') },
      h('span', { class: 'dk-ico' }, icon(iconName)),
      h('div', { class: 'dk-body' }, h('div', { class: 'dk-t' }, title), h('div', { class: 'dk-v' }, value, unit ? h('small', null, unit) : null), note ? h('div', { class: 'dk-note' }, note) : null));
  }
  function panel(title, more, body, cls) {
    return h('section', { class: 'dp' + (cls ? ' ' + cls : '') },
      h('div', { class: 'dp-head' }, h('h2', null, title), more ? h('a', { class: 'dp-more', href: more[0] }, more[1] || '더보기', icon('chev')) : null),
      body);
  }
  // 가로 막대 — 글자는 HTML 이라 좁은 화면에서도 작아지지 않습니다
  function hbars(items, opt) {
    var max = Math.max.apply(null, items.map(function (x) { return x.value; }).concat([1]));
    return h('ul', { class: 'hbars' + (opt && opt.cls ? ' ' + opt.cls : ''), 'aria-label': opt && opt.label },
      items.map(function (x) {
        var name = x.href ? h('a', { href: x.href, class: 'hb-k' }, x.key) : h('span', { class: 'hb-k' }, x.key);
        return h('li', { class: 'hb' + (x.cls ? ' ' + x.cls : '') }, name,
          h('span', { class: 'hb-track' }, h('span', { class: 'hb-fill', style: 'width:' + (x.value ? Math.max(2, x.value / max * 100) : 0) + '%' })),
          h('span', { class: 'hb-v' }, fmt(x.value) + (x.unit || '건'), x.sub ? h('small', null, ' ' + x.sub) : null));
      }));
  }
  function vbars(items, opt) {
    var max = Math.max.apply(null, items.map(function (x) { return x.value; }).concat([1]));
    return h('div', { class: 'vbars', role: 'img', 'aria-label': opt.label },
      items.map(function (x) {
        return h('div', { class: 'vbar', title: x.tip || null },
          h('span', { class: 'vbar-v' }, fmt(x.value)),
          h('span', { class: 'vbar-track' }, h('span', { class: 'vbar-fill', style: 'height:' + (x.value ? Math.max(3, x.value / max * 100) : 0) + '%' })),
          h('span', { class: 'vbar-k' }, x.key));
      }));
  }
  function monthBars(months, label) {
    var list = months.slice(-12);
    return vbars(list.map(function (m) { return { key: +m.month.slice(5) + '월', value: m.count, tip: m.month + ' · ' + m.count + '건' }; }), { label: label });
  }
  function paretoBars(byType) {
    return hbars(byType.map(function (t) { return { key: t.type, value: t.count, sub: '(' + t.pct + '% · 누적 ' + t.cum + '%)', cls: t.type === L.UNCLASSIFIED ? 'unc' : '' }; }), { label: '불량유형별 건수 (파레토)', cls: 'wide' });
  }
  function stackBar(st) {
    var tot = st.total || 1;
    return h('span', { class: 'stack', role: 'img', 'aria-label': '완료 ' + st.done + ', 미완료 ' + st.open + ', 보류 ' + st.hold },
      ['done', 'open', 'hold'].map(function (k) { return st[k] ? h('span', { class: 'stack-' + k, style: 'width:' + (st[k] / tot * 100) + '%' }) : null; }));
  }

  // ── 라우팅 ────────────────────────────────────────────────
  function routeParts() { return location.hash.replace(/^#\/?/, '').split('/'); }
  function route() { return routeParts()[0] || 'home'; }
  function routeArg() { var p = routeParts()[1]; try { return p ? decodeURIComponent(p) : ''; } catch (e) { return ''; } }
  function render() {
    var r = route();
    var main = document.getElementById('main');
    main.textContent = '';
    document.querySelectorAll('.sidebar a[data-route]').forEach(function (a) {
      if (a.getAttribute('data-route') === r) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    document.getElementById('sampleBanner').hidden = !db._sample;
    if (!S.available()) document.getElementById('storeBanner').hidden = false;
    document.body.setAttribute('data-route', r);
    if (r === 'upload') renderUpload(main);
    else if (r === 'analysis') renderAnalysis(main);
    else if (r === 'plan') renderPlan(main);
    else if (r === 'progress') renderProgress(main);
    else if (r === 'report') renderReport(main);
    else if (r === 'settings') renderSettings(main);
    else renderHome(main);
  }
  window.addEventListener('hashchange', function () { setMenu(false); closeDrawer(true); render(); window.scrollTo(0, 0); });

  // 좁은 화면(1024px 미만)에서는 왼쪽 메뉴를 「메뉴」 단추로 여닫습니다.
  var menuBtn = document.getElementById('menuBtn'), sidebar = document.getElementById('sidebar'), backdrop = document.getElementById('sideBackdrop');
  function setMenu(open) {
    document.body.classList.toggle('menu-open', open);
    menuBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    menuBtn.setAttribute('aria-label', open ? '메뉴 닫기' : '메뉴 열기');
    backdrop.hidden = !open;
    if (open) { var first = sidebar.querySelector('a'); if (first) first.focus(); }
  }
  menuBtn.addEventListener('click', function () { setMenu(!document.body.classList.contains('menu-open')); });
  backdrop.addEventListener('click', function () { setMenu(false); menuBtn.focus(); });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    if (!document.getElementById('drawer').hidden) { closeDrawer(); return; }
    if (document.body.classList.contains('menu-open')) { setMenu(false); menuBtn.focus(); }
  });
  sidebar.addEventListener('click', function (e) { if (e.target.closest('a')) setMenu(false); });
  var WEEK = ['일', '월', '화', '수', '목', '금', '토'];
  function nowText() {
    var d = new Date();
    function p(n) { return (n < 10 ? '0' : '') + n; }
    document.getElementById('nowText').textContent = d.getFullYear() + '. ' + p(d.getMonth() + 1) + '. ' + p(d.getDate()) + ' (' + WEEK[d.getDay()] + ') ' + p(d.getHours()) + ':' + p(d.getMinutes());
  }
  nowText(); setInterval(nowText, 30000);
  // 간트는 700px 을 사이에 두고 칸 구성이 바뀌므로 그 경계를 넘을 때만 다시 그립니다
  var wasNarrow = document.documentElement.clientWidth < 700;
  window.addEventListener('resize', function () {
    clearTimeout(render._r);
    render._r = setTimeout(function () {
      var n = document.documentElement.clientWidth < 700;
      if (n !== wasNarrow) { wasNarrow = n; if (route() === 'plan' || route() === 'progress') render(); }
    }, 200);
  });

  // ── 예시 데이터 ───────────────────────────────────────────
  function sampleButton(primary) {
    return h('button', { type: 'button', class: 'btn' + (primary ? ' btn-primary' : ''), onclick: loadSample }, '예시 데이터로 시작');
  }
  function loadSample() {
    function go() {
      var keep = { types: db.types, depts: db.depts, rules: db.rules, mask: db.mask, mapping: db.mapping };
      db = L.emptyDb();
      Object.assign(db, keep);
      var m = L.mergeRecords([], Sample.records(), 0);
      db.records = m.records; db.seq = m.seq; db._sample = true;
      db.aliases = L.updateAliases(db.aliases, db.records);
      db.files = [{ name: '예시 데이터(가상)', at: nowIso(), added: m.added, dup: m.dup }];
      save();
      toast('예시 데이터 ' + m.added + '건을 넣었습니다(협력사 A~F, 가상).');
      if (route() !== 'analysis') location.hash = '#/analysis'; else render();
    }
    if (db.records.length && !db._sample) confirmBox('예시 데이터로 바꾸기', '지금 있는 Raw data ' + db.records.length + '건과 Master Plan 을 지우고 예시 데이터를 넣습니다.\n필요하면 먼저 설정에서 JSON 백업을 받아 두세요.', '바꾸기', go, true);
    else go();
  }

  // ══ 대시보드 ═════════════════════════════════════════════
  function renderHome(main) {
    var rows = all();
    var b = L.dateBounds(rows);
    main.appendChild(h('section', { class: 'hero' },
      h('div', { class: 'hero-text' },
        h('h1', null, '협력사 품질 현황'),
        h('p', null, '협력사 공정불량을 유형별로 나누어 보고, 개선이 필요한 항목부터 Master Plan 으로 관리합니다.')),
      b.min ? h('span', { class: 'hero-chip' }, icon('cal'), b.min + ' ~ ' + b.max) : null));
    if (!rows.length) {
      main.appendChild(h('section', { class: 'dp dash-empty' },
        h('h2', null, '아직 공정불량 Raw data 가 없습니다'),
        h('p', null, '발생일자·품명·품번·문제점·협력사 열이 있는 Excel 을 올리면 이 화면의 지표·차트가 채워집니다. 실제 파일이 아직 없으면 예시 데이터(가상)로 흐름을 볼 수 있습니다.'),
        h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/upload' }, 'RAW DATA 업로드'), sampleButton(false))));
      main.appendChild(flowGuide());
      return;
    }
    var an = L.analyze(rows, {}, db.rules);
    var allPlan = [];
    Object.keys(db.plans).forEach(function (s) { allPlan = allPlan.concat(db.plans[s].rows); });
    var ps = L.planStats(allPlan, today());
    var last3 = L.analyze(rows, { from: L.monthAdd(L.monthOf(an.to), -2) + '-01', to: an.to }, db.rules);
    main.appendChild(h('section', { class: 'dkpis' },
      kpi('불량 발생 건수', 'doc', fmt(an.total), '건', '최근 3개월 ' + last3.total + '건'),
      kpi('협력사', 'factory', an.suppliers.length, '곳', an.unclassified ? '유형 미분류 ' + an.unclassified + '건' : '모든 건 유형 분류됨'),
      kpi('개선 필요 항목', 'alert', an.issues.length, '건', '반복 · 증가 추세 · 반복 품번 규칙', 'warn'),
      kpi('개선활동 완료', 'check', ps.done + ' / ' + ps.total, '', ps.total ? '완료율 ' + ps.rate + '% · 지연 ' + ps.overdue + '건' : 'Master Plan 을 아직 만들지 않았습니다', 'ok')));

    main.appendChild(h('div', { class: 'dcharts' },
      panel('협력사별 불량 건수', ['#/analysis', '불량 분석'], hbars(an.suppliers.map(function (s) { return { key: s.name, value: s.count, href: '#/analysis/' + encodeURIComponent(s.name) }; }), { label: '협력사별 불량 건수' })),
      panel('불량유형 파레토', ['#/analysis', '불량 분석'], paretoBars(an.byType)),
      panel('월별 불량 추이', null, h('div', null, h('p', { class: 'dp-sub' }, '최근 12개월 · 막대 = 건수'), monthBars(an.months, '월별 불량 건수')))));

    var tb = h('tbody');
    an.issues.slice(0, 8).forEach(function (i) {
      tb.appendChild(h('tr', null,
        h('td', { class: 'nw' }, i.supplier), h('td', { class: 'nw' }, i.type), h('td', { class: 'num' }, i.count + '건'),
        h('td', null, issueFlags(i)),
        h('td', null, h('a', { class: 'btn btn-sm2', href: '#/plan/' + encodeURIComponent(i.supplier) }, '계획'))));
    });
    var issues = panel('개선 필요 항목 TOP 8', ['#/analysis', '전체 보기'], an.issues.length
      ? h('div', { class: 'table-wrap flat' }, h('table', { class: 'list dash-table' }, h('thead', null, h('tr', null, ['협력사', '불량유형', '건수', '근거', 'Master Plan'].map(function (t) { return h('th', null, t); }))), tb))
      : h('p', { class: 'note' }, '지금 규칙으로 걸린 항목이 없습니다. 설정에서 기준을 확인하세요.'), 'a-issues');

    var pt = h('tbody');
    suppliers().forEach(function (s) {
      var p = db.plans[s]; if (!p || !p.rows.length) return;
      var st = L.planStats(p.rows, today());
      pt.appendChild(h('tr', { class: 'clickable', tabindex: '0', onclick: function () { location.hash = '#/plan/' + encodeURIComponent(s); }, onkeydown: function (e) { if (e.key === 'Enter') location.hash = '#/plan/' + encodeURIComponent(s); } },
        h('td', { class: 'nw' }, s), h('td', { class: 'stack-cell' }, stackBar(st)), h('td', { class: 'num' }, st.done + '/' + st.total), h('td', { class: 'num' }, st.overdue ? h('span', { class: 'late' }, st.overdue + '건') : '0')));
    });
    var prog = panel('협력사별 추진현황', ['#/progress', '추진현황'], pt.children.length
      ? h('div', null, h('div', { class: 'table-wrap flat' }, h('table', { class: 'list dash-table' }, h('thead', null, h('tr', null, ['협력사', '상태', '완료', '지연'].map(function (t) { return h('th', null, t); }))), pt)), legend())
      : h('p', { class: 'note' }, '아직 Master Plan 이 없습니다. 「MASTER PLAN」에서 협력사를 골라 「분석 및 초안 생성」을 누르세요.'), 'a-prog');
    main.appendChild(h('div', { class: 'dash-low' }, issues, prog));
    main.appendChild(flowGuide());
  }
  function issueFlags(i) {
    return h('span', { class: 'flags' },
      i.flags.repeat ? h('span', { class: 'flag', title: i.reasons.filter(function (r) { return r.indexOf('반복 발생') === 0; })[0] }, '반복') : null,
      i.flags.trend ? h('span', { class: 'flag up', title: i.reasons.filter(function (r) { return r.indexOf('증가') === 0; })[0] }, '증가 ' + i.prev + '→' + i.recent) : null,
      i.flags.part ? h('span', { class: 'flag part', title: i.reasons.filter(function (r) { return r.indexOf('반복 품번') === 0; })[0] }, '품번 ' + i.repeatParts[0].partNo + ' ' + i.repeatParts[0].count + '건') : null);
  }
  function legend() {
    return h('ul', { class: 'legend' }, ['done', 'open', 'hold'].map(function (k) { return h('li', null, h('i', { class: 'lg-' + k }), L.STATUS[k]); }));
  }
  function flowGuide() {
    var steps = [
      ['#/upload', 'RAW DATA 업로드', 'Excel 을 올리면 열을 자동으로 맞춥니다'],
      ['#/analysis', '불량 분석', '문제점 글로 불량유형을 나누고 개선 필요 항목을 찾습니다'],
      ['#/plan', 'MASTER PLAN', '유형별 개선대책 초안 → AI 로 구체화 → 간트로 계획 확정'],
      ['#/progress', '추진현황', '담당자가 실적과 완료·미완료·보류 상태를 적습니다'],
      ['#/report', '보고서 출력', '확정한 Master Plan 을 Excel 보고서로 내보냅니다']
    ];
    return panel('일하는 순서', null, h('ol', { class: 'flow' }, steps.map(function (s, i) {
      return h('li', null, h('a', { href: s[0] }, h('span', { class: 'flow-n' }, String(i + 1)), h('span', { class: 'flow-t' }, h('strong', null, s[1]), h('span', null, s[2])), icon('chev', 'q-chev')));
    })), 'flow-panel');
  }

  // ══ RAW DATA 업로드 ═══════════════════════════════════════
  function decodeCsv(buf) {
    var u8 = new Uint8Array(buf), t;
    try { t = new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch (e) { try { t = new TextDecoder('euc-kr').decode(u8); } catch (e2) { t = new TextDecoder('utf-8').decode(u8); } }
    return t.replace(/^﻿/, '');
  }
  function readOne(f) {
    return f.arrayBuffer().then(function (buf) {
      var wb = /\.csv$/i.test(f.name) ? XLSX.read(decodeCsv(buf), { type: 'string' }) : XLSX.read(buf, { type: 'array' });
      var sheets = wb.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }; });
      var pick = L.pickSheet(sheets) || sheets[0];
      var st = { name: f.name, sheets: sheets, sheet: pick ? pick.name : '', ov: null };
      var det = L.detectColumns(pick ? pick.aoa : []);
      var sig = det.headers.map(L.normHead).join('|');
      if (db.mapping && db.mapping[sig]) st.ov = Object.assign({}, db.mapping[sig]);
      return st;
    }).catch(function (e) { return { name: f.name, error: '읽지 못했습니다 (' + (e && e.message || e) + ')' }; });
  }
  function readFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;
    toast(files.length + '개 파일을 읽는 중…');
    Promise.all(files.map(readOne)).then(function (list) { ui.imp = list; render(); });
  }
  function parsedOf(st) {
    var sh = st.sheets.filter(function (s) { return s.name === st.sheet; })[0];
    if (!sh) return null;
    return { aoa: sh.aoa, r: L.parseSheet(sh.aoa, st.ov ? { overrides: st.ov } : null) };
  }
  function templateButton() {
    return h('button', { type: 'button', class: 'btn', onclick: function () {
      writeXlsx('공정불량_RawData_양식.xlsx', { 'RAW': [['발생일자', '품명', '품번', '문제점', '협력사', '수량'], ['2026-10-05', '(예) 브래킷', '(예) TP-0001', '(예) 용접부 균열 발생', '(예) 협력사 A', 1]] }, { 'RAW': [12, 14, 14, 36, 14, 8] });
    } }, '빈 양식 받기');
  }
  function renderUpload(main) {
    main.appendChild(pageHead('RAW DATA 업로드', '협력사 공정불량 Excel(또는 CSV)을 올립니다. 파일은 서버로 보내지 않고 이 브라우저에서만 읽습니다.'));
    var input = h('input', { type: 'file', accept: '.xlsx,.xls,.csv', multiple: true, onchange: function (e) { readFiles(e.target.files); e.target.value = ''; } });
    var drop = h('label', { class: 'drop' }, icon('down', 'drop-ico'), h('strong', null, 'Excel 파일 고르기'), h('span', { class: 'note' }, '.xlsx · .xls · .csv — 여러 개를 함께 골라도 됩니다(여기로 끌어 놓아도 됨)'), input);
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); readFiles(e.dataTransfer.files); });
    main.appendChild(h('section', { class: 'card' },
      h('h2', null, '1. 파일 올리기'),
      h('p', null, '필요한 열: ', h('strong', null, '발생일자 · 품명 · 품번 · 문제점 · 협력사'), ' (수량은 있으면 함께 읽습니다). 열 이름이 조금 달라도(업체명·불량내용·P/N 등) 자동으로 찾고, 위에 제목 줄이 있어도 머리행을 찾아 읽습니다.'),
      drop,
      h('div', { class: 'btn-row', style: 'margin-top:12px' }, templateButton(), sampleButton(false)),
      h('p', { class: 'note', style: 'margin-top:10px' }, '시험용 가상 파일: 리포의 samples/예시_협력사_공정불량_RawData.xlsx')));

    if (ui.imp) main.appendChild(importCard());

    var rows = db.records;
    var b = L.dateBounds(rows);
    var cur = h('section', { class: 'card' }, h('h2', null, '지금 들어 있는 데이터'));
    if (!rows.length) cur.appendChild(h('p', { class: 'note' }, '아직 없습니다.'));
    else {
      cur.appendChild(h('div', { class: 'kpis' },
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '건수'), h('div', { class: 'v' }, fmt(rows.length), h('small', null, '건'))),
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '협력사'), h('div', { class: 'v' }, suppliers().length, h('small', null, '곳'))),
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '기간'), h('div', { class: 'v sm' }, b.min + ' ~ ' + b.max)),
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '저장 크기'), h('div', { class: 'v sm' }, S.sizeKb() + ' KB'))));
      if (db.files.length) cur.appendChild(h('details', { class: 'files' }, h('summary', null, '가져온 기록 ' + db.files.length + '건'),
        h('ul', null, db.files.slice().reverse().map(function (f) { return h('li', null, f.at.replace('T', ' ').slice(0, 16) + ' · ' + f.name + ' · 새로 ' + f.added + '건' + (f.dup ? ', 같은 행 ' + f.dup + '건 건너뜀' : '')); }))));
      var recent = rows.slice().sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; }).slice(0, 8);
      cur.appendChild(h('h3', { style: 'margin-top:14px' }, '최근 8건'));
      cur.appendChild(rawTable(recent));
      cur.appendChild(h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn', onclick: exportRaw }, 'Raw data Excel 로 내보내기 (불량유형 포함)'),
        h('a', { class: 'btn btn-primary', href: '#/analysis' }, '불량 분석으로'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: clearAll }, '전부 지우기')));
    }
    main.appendChild(cur);
  }
  function rawTable(rows) {
    return h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
      h('thead', null, h('tr', null, ['발생일자', '협력사', '품번', '품명', '문제점'].map(function (t) { return h('th', null, t); }))),
      h('tbody', null, rows.map(function (r) { return h('tr', null, h('td', { class: 'nw' }, r.date || '(날짜 없음)'), h('td', { class: 'nw' }, r.supplier), h('td', { class: 'nw' }, r.partNo), h('td', null, r.partName), h('td', { class: 'wide' }, r.problem)); }))));
  }
  function importCard() {
    var card = h('section', { class: 'card import' }, h('h2', null, '2. 열 연결 확인'));
    var ok = 0, total = 0;
    ui.imp.forEach(function (st, fi) {
      var box = h('div', { class: 'imp-file' }, h('h3', null, st.name));
      card.appendChild(box);
      if (st.error) { box.appendChild(h('p', { class: 'alert error' }, st.error)); return; }
      if (st.sheets.length > 1) {
        var ss = select(st.sheets.map(function (s) { return s.name; }), st.sheet, { 'aria-label': '시트' });
        ss.addEventListener('change', function () { st.sheet = ss.value; st.ov = null; render(); });
        box.appendChild(field('시트', ss, '머리행이 가장 잘 맞는 시트를 골라 두었습니다.', 'inline'));
      }
      var p = parsedOf(st);
      if (!p) return;
      var det = p.r.det;
      if (det.headerRow < 0) {
        box.appendChild(h('p', { class: 'alert warn' }, '머리행(발생일자·품번·협력사 같은 열 이름 줄)을 찾지 못했습니다. 첫 줄을 머리행으로 보고 열을 직접 골라 주세요.'));
        if (!st.ov) { st.ov = {}; }
        p = { aoa: p.aoa, r: L.parseSheet(p.aoa, { overrides: st.ov, headerRow: 0 }) };
        det = p.r.det;
      }
      var headers = det.headers;
      var grid = h('div', { class: 'map-grid' });
      L.FIELDS.forEach(function (f) {
        var opts = [['-1', '(연결 안 함)']].concat(headers.map(function (x, i) { return [String(i), String.fromCharCode(65 + (i % 26)) + (i >= 26 ? Math.floor(i / 26) : '') + ': ' + (x || '(빈 칸)')]; }));
        var s = select(opts, det.map[f.id] == null ? '-1' : String(det.map[f.id]), { 'aria-label': f.label });
        s.addEventListener('change', function () { st.ov = Object.assign({}, st.ov || {}); st.ov[f.id] = +s.value; render(); });
        grid.appendChild(field(f.label + (f.req ? ' *' : ''), s, null, det.map[f.id] == null && f.req ? 'miss' : ''));
      });
      box.appendChild(grid);
      total++;
      if (det.missing.length) {
        box.appendChild(h('p', { class: 'alert warn' }, '꼭 있어야 하는 열을 연결해 주세요: ' + det.missing.map(function (id) { return L.FIELDS.filter(function (f) { return f.id === id; })[0].label; }).join(', ')));
        return;
      }
      ok++;
      var r = p.r;
      box.appendChild(h('p', { class: 'note' }, '읽은 행 ' + r.records.length + '건' + (r.skipped ? ' · 문제점·협력사가 비어 건너뛴 행 ' + r.skipped + '건' : '') + (r.badDate ? ' · 날짜를 못 읽은 행 ' + r.badDate + '건(날짜 없이 넣고 월별 집계에서는 빠짐)' : '')));
      box.appendChild(rawTable(r.records.slice(0, 5)));
    });
    var modeName = 'impmode';
    var modes = h('div', { class: 'radios' },
      h('label', { class: 'check' }, h('input', { type: 'radio', name: modeName, checked: ui.impMode === 'add', onchange: function () { ui.impMode = 'add'; } }), '지금 데이터에 더하기 (같은 행은 건너뜀)'),
      h('label', { class: 'check' }, h('input', { type: 'radio', name: modeName, checked: ui.impMode === 'replace', onchange: function () { ui.impMode = 'replace'; } }), '지금 데이터를 지우고 바꾸기'));
    card.appendChild(h('h2', { style: 'margin-top:16px' }, '3. 가져오기'));
    if (db._sample) card.appendChild(h('p', { class: 'alert info' }, '지금은 예시 데이터라, 가져오면 예시 데이터를 지우고 바꿉니다.'));
    else if (db.records.length) card.appendChild(modes);
    card.appendChild(h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn btn-primary', disabled: !ok, onclick: doImport }, ok + '/' + total + '개 파일 가져오기'),
      h('button', { type: 'button', class: 'btn', onclick: function () { ui.imp = null; render(); } }, '그만두기')));
    return card;
  }
  function doImport() {
    var recs = [], names = [];
    db.mapping = db.mapping || {};
    ui.imp.forEach(function (st) {
      if (st.error) return;
      var p = parsedOf(st); if (!p) return;
      var r = st.ov ? L.parseSheet(p.aoa, { overrides: st.ov, headerRow: p.r.det.headerRow < 0 ? 0 : undefined }) : p.r;
      if (r.det.missing.length) return;
      if (st.ov) db.mapping[r.det.headers.map(L.normHead).join('|')] = st.ov;
      recs = recs.concat(r.records); names.push(st.name);
    });
    var replace = db._sample || ui.impMode === 'replace' || !db.records.length;
    if (replace) { db.records = []; db.seq = 0; if (db._sample) { db.plans = {}; db.files = []; } delete db._sample; }
    var m = L.mergeRecords(db.records, recs, db.seq);
    db.records = m.records; db.seq = m.seq;
    db.aliases = L.updateAliases(db.aliases, db.records);
    db.files.push({ name: names.join(', '), at: nowIso(), added: m.added, dup: m.dup });
    save();
    ui.imp = null;
    toast('새로 ' + m.added + '건을 넣었습니다' + (m.dup ? ' (같은 행 ' + m.dup + '건 건너뜀)' : '') + '.');
    location.hash = '#/analysis';
  }
  function exportRaw() {
    var rows = all();
    var aoa = [['발생일자', '품명', '품번', '문제점', '협력사', '수량', '불량유형', '분류 근거']];
    rows.forEach(function (r) { aoa.push([r.date, r.partName, r.partNo, r.problem, r.supplier, r.qty == null ? '' : r.qty, r.type, r.manual ? '직접 지정' : r.hits.join(', ')]); });
    writeXlsx('공정불량_RawData_분류' + tag() + '_' + today() + '.xlsx', { 'RAW': aoa }, { 'RAW': [12, 14, 14, 36, 14, 8, 14, 18] });
  }
  function clearAll() {
    confirmBox('전부 지우기', 'Raw data ' + db.records.length + '건과 모든 협력사의 Master Plan 을 지웁니다. 불량유형 사전·규칙·AI 설정은 남습니다.\n되돌릴 수 없으니 필요하면 먼저 설정에서 JSON 백업을 받아 두세요.', '지우기', function () {
      var keep = { types: db.types, depts: db.depts, rules: db.rules, mask: db.mask, mapping: db.mapping };
      db = L.emptyDb(); Object.assign(db, keep); save(); toast('지웠습니다.'); render();
    }, true);
  }

  // ══ 불량 분석 ═════════════════════════════════════════════
  function periodControls(state, onChange, bounds) {
    var sup = select([['', '전체 협력사']].concat(suppliers().map(function (s) { return [s, s]; })), state.supplier, { 'aria-label': '협력사' });
    sup.addEventListener('change', function () { state.supplier = sup.value; state.page = 0; onChange(); });
    var f = dateInput(state.from, { 'aria-label': '시작일', min: bounds.min, max: bounds.max });
    var t = dateInput(state.to, { 'aria-label': '종료일', min: bounds.min, max: bounds.max });
    f.addEventListener('change', function () { state.from = f.value; state.page = 0; onChange(); });
    t.addEventListener('change', function () { state.to = t.value; state.page = 0; onChange(); });
    function quick(label, months) {
      return h('button', { type: 'button', class: 'chip-btn', onclick: function () {
        state.to = ''; state.from = months ? L.monthAdd(L.monthOf(bounds.max), -(months - 1)) + '-01' : ''; state.page = 0; onChange();
      } }, label);
    }
    return h('div', { class: 'filters' },
      field('협력사', sup), field('분석 기간 — 시작', f), field('끝', t),
      h('div', { class: 'field' }, h('span', null, '빠른 기간'), h('div', { class: 'chips' }, quick('최근 3개월', 3), quick('최근 6개월', 6), quick('전체', 0))));
  }
  function renderAnalysis(main) {
    var arg = routeArg();
    if (arg && suppliers().indexOf(arg) >= 0) { ui.an.supplier = arg; history.replaceState(null, '', '#/analysis'); }
    if (!db.records.length) { emptyNotice(main, '불량 분석'); return; }
    var rows = all(), st = ui.an, b = L.dateBounds(rows);
    var an = L.analyze(rows, { supplier: st.supplier, from: st.from, to: st.to }, db.rules);
    main.appendChild(pageHead('불량 분석', '문제점 글의 키워드로 불량유형을 나누고, 협력사별 발생 현황과 개선이 필요한 항목을 찾습니다.'));
    main.appendChild(h('section', { class: 'card' }, periodControls(st, render, b)));
    main.appendChild(h('section', { class: 'dkpis' },
      kpi('불량 발생 건수', 'doc', fmt(an.total), '건', an.from ? an.from + ' ~ ' + an.to : ''),
      an.hasQty ? kpi('불량 수량', 'layers', fmt(an.qty), '개', '수량 열이 있는 행만 합계') : kpi('협력사', 'factory', an.suppliers.length, '곳'),
      kpi('개선 필요 항목', 'alert', an.issues.length, '건', '설정의 규칙 기준', 'warn'),
      kpi('유형 미분류', 'list', an.unclassified, '건', an.unclassified ? '아래 표에서 유형을 고르거나 키워드를 더하세요' : '모두 분류됨', an.unclassified ? 'warn' : 'ok')));
    if (!an.total) { main.appendChild(h('section', { class: 'card' }, h('p', null, '이 조건에 맞는 공정불량이 없습니다.'))); return; }

    var charts = [
      panel('불량유형 파레토', null, paretoBars(an.byType)),
      panel('월별 추이', null, h('div', null, h('p', { class: 'dp-sub' }, '최근 12개월까지 · 막대 = 건수'), monthBars(an.months, '월별 불량 건수')))
    ];
    if (!st.supplier) charts.unshift(panel('협력사별 건수', null, hbars(an.suppliers.map(function (s) { return { key: s.name, value: s.count, href: '#/analysis/' + encodeURIComponent(s.name) }; }), { label: '협력사별 불량 건수' })));
    main.appendChild(h('div', { class: st.supplier ? 'grid-2 gap-b' : 'dcharts' }, charts));
    if (!st.supplier && an.suppliers.length > 1) main.appendChild(crossTable(an));

    // 개선 필요 항목
    var list = h('ul', { class: 'issues' });
    an.issues.forEach(function (i) {
      list.appendChild(h('li', { class: 'issue' },
        h('div', { class: 'issue-head' }, h('strong', null, i.supplier + ' · ' + i.type), h('span', { class: 'num' }, i.count + '건'), issueFlags(i),
          h('a', { class: 'btn btn-sm2', href: '#/plan/' + encodeURIComponent(i.supplier) }, 'Master Plan')),
        h('ul', { class: 'reasons' }, i.reasons.map(function (r) { return h('li', null, r); })),
        i.problems.length ? h('p', { class: 'note' }, '문제점 예: ' + i.problems.slice(0, 3).map(function (p) { return '「' + p + '」'; }).join(' ')) : null));
    });
    var r = db.rules;
    main.appendChild(h('section', { class: 'card' }, h('div', { class: 'card-head' }, h('h2', null, '개선 필요 항목 ' + an.issues.length + '건'), h('a', { class: 'dp-more', href: '#/settings' }, '기준 바꾸기', icon('chev'))),
      h('p', { class: 'note' }, '협력사 × 불량유형마다 — 반복 발생(기간 안 ' + r.repeatMin + '건 이상) · 증가 추세(최근 ' + r.trendMonths + '개월 ' + r.trendMin + '건 이상이고 직전 ' + r.trendMonths + '개월보다 ' + r.trendRate + '% 이상 많음) · 반복 품번(같은 품번 ' + r.partRepeatMin + '건 이상) 가운데 하나라도 걸리면 올립니다. 위에 있을수록 건수·증가·반복 품번이 많은 항목입니다.'),
      an.issues.length ? list : h('p', null, '규칙에 걸린 항목이 없습니다.')));

    // 반복 품번
    if (an.repeatParts.length) main.appendChild(h('section', { class: 'card' }, h('h2', null, '반복 품번 ' + an.repeatParts.length + '개'),
      h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
        h('thead', null, h('tr', null, ['협력사', '품번', '품명', '건수', '불량유형', '최근 발생'].map(function (t) { return h('th', null, t); }))),
        h('tbody', null, an.repeatParts.slice(0, 30).map(function (p) { return h('tr', null, h('td', { class: 'nw' }, p.supplier), h('td', { class: 'nw' }, p.partNo), h('td', null, p.partName), h('td', { class: 'num' }, p.count), h('td', null, p.types.join(', ')), h('td', { class: 'nw' }, p.last)); }))))));

    main.appendChild(classifyCard(an));
  }
  function crossTable(an) {
    var types = an.byType.map(function (t) { return t.type; });
    var m = {}, max = 1;
    an.rows.forEach(function (r) { var k = r.supplier + '\u0000' + r.type; m[k] = (m[k] || 0) + 1; max = Math.max(max, m[k]); });
    var issueKey = {}; an.issues.forEach(function (i) { issueKey[i.supplier + '\u0000' + i.type] = 1; });
    return h('section', { class: 'card' }, h('h2', null, '협력사 × 불량유형'),
      h('p', { class: 'note' }, '칸이 진할수록 건수가 많습니다. 굵은 테두리는 개선 필요 항목입니다.'),
      h('div', { class: 'table-wrap' }, h('table', { class: 'list cross' },
        h('thead', null, h('tr', null, h('th', null, '협력사'), types.map(function (t) { return h('th', null, t); }), h('th', { class: 'num' }, '합계'))),
        h('tbody', null, an.suppliers.map(function (s) {
          return h('tr', null, h('th', { scope: 'row' }, h('a', { href: '#/analysis/' + encodeURIComponent(s.name) }, s.name)),
            types.map(function (t) {
              var v = m[s.name + '\u0000' + t] || 0, a = v / max;
              return h('td', { class: 'num heat' + (issueKey[s.name + '\u0000' + t] ? ' hot' : ''), style: v ? 'background:rgba(31,95,168,' + (0.08 + a * 0.62).toFixed(2) + ');color:' + (a > 0.55 ? '#fff' : 'inherit') : null }, v || '');
            }), h('td', { class: 'num' }, s.count));
        })))));
  }
  function classifyCard(an) {
    var st = ui.an;
    var names = L.typeNames(db.types);
    var rows = an.rows.filter(function (r) {
      if (st.onlyUnc && r.type !== L.UNCLASSIFIED) return false;
      if (st.type && r.type !== st.type) return false;
      if (st.q) { var q = st.q.toLowerCase(); if ((r.problem + ' ' + r.partNo + ' ' + r.partName).toLowerCase().indexOf(q) < 0) return false; }
      return true;
    }).sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
    var pages = Math.max(1, Math.ceil(rows.length / PAGE));
    if (st.page >= pages) st.page = pages - 1;
    var tsel = select([['', '전체 유형']].concat(names.concat([L.UNCLASSIFIED]).map(function (n) { return [n, n]; })), st.type, { 'aria-label': '불량유형' });
    tsel.addEventListener('change', function () { st.type = tsel.value; st.page = 0; render(); });
    var q = h('input', { type: 'search', value: st.q, placeholder: '문제점·품번 찾기', 'aria-label': '문제점·품번 찾기' });
    q.addEventListener('change', function () { st.q = q.value.trim(); st.page = 0; render(); });
    var unc = h('input', { type: 'checkbox', checked: st.onlyUnc, onchange: function (e) { st.onlyUnc = e.target.checked; st.page = 0; render(); } });
    var tb = h('tbody');
    rows.slice(st.page * PAGE, (st.page + 1) * PAGE).forEach(function (r) {
      var raw = db.records.filter(function (x) { return x.id === r.id; })[0];
      var s = select([['', '(자동) ' + (raw && raw.typeManual ? L.classify(r.problem, db.types).type : r.type)]].concat(names.concat([L.UNCLASSIFIED]).map(function (n) { return [n, n]; })), raw && raw.typeManual || '', { 'aria-label': '불량유형 고르기', class: 'type-sel' + (r.manual ? ' manual' : '') });
      s.addEventListener('change', function () { if (!raw) return; if (s.value) raw.typeManual = s.value; else delete raw.typeManual; save(); render(); });
      tb.appendChild(h('tr', { class: r.type === L.UNCLASSIFIED ? 'row-warn' : '' },
        h('td', { class: 'nw' }, r.date), h('td', { class: 'nw' }, r.supplier), h('td', { class: 'nw' }, r.partNo), h('td', null, r.partName), h('td', { class: 'wide' }, r.problem),
        h('td', null, s), h('td', { class: 'note' }, r.manual ? '직접 지정' : r.hits.join(', '))));
    });
    var kw = h('input', { placeholder: '예: 미첨부', 'aria-label': '더할 키워드' });
    var kwType = select(names, names[0], { 'aria-label': '키워드를 더할 유형' });
    var addKw = h('form', { class: 'kw-add', onsubmit: function (e) {
      e.preventDefault();
      var k = kw.value.trim(); if (!k) return;
      var t = db.types.filter(function (x) { return x.name === kwType.value; })[0];
      if (t.keywords.indexOf(k) < 0) t.keywords.push(k);
      save(); toast('「' + t.name + '」 키워드에 「' + k + '」를 더했습니다.'); render();
    } }, field('키워드', kw), field('유형', kwType), h('button', { type: 'submit', class: 'btn' }, '키워드 더하기'));
    return h('section', { class: 'card' },
      h('div', { class: 'card-head' }, h('h2', null, '분류 결과 ' + rows.length + '건'), h('a', { class: 'dp-more', href: '#/settings' }, '키워드 사전', icon('chev'))),
      h('p', { class: 'note' }, '불량유형은 문제점 글에 든 키워드로 정합니다(근거 칸). 틀리면 그 행에서 유형을 직접 고르세요 — 직접 고른 값이 먼저입니다. 같은 낱말이 자주 나오면 키워드로 더하는 편이 낫습니다.'),
      h('div', { class: 'filters' }, field('유형', tsel), field('찾기', q), h('label', { class: 'check' }, unc, '미분류만')),
      h('div', { class: 'table-wrap', style: 'margin-top:12px' }, h('table', { class: 'list' },
        h('thead', null, h('tr', null, ['발생일자', '협력사', '품번', '품명', '문제점', '불량유형', '근거'].map(function (t) { return h('th', null, t); }))), tb)),
      pager(st, pages, rows.length),
      h('details', { class: 'kw-box' }, h('summary', null, '키워드 바로 더하기'), addKw));
  }
  function pager(st, pages, n) {
    if (pages <= 1) return null;
    return h('div', { class: 'pager' },
      h('button', { type: 'button', class: 'btn', disabled: st.page === 0, onclick: function () { st.page--; render(); } }, '이전'),
      h('span', null, (st.page + 1) + ' / ' + pages + ' 쪽 · ' + n + '건'),
      h('button', { type: 'button', class: 'btn', disabled: st.page >= pages - 1, onclick: function () { st.page++; render(); } }, '다음'));
  }

  // ══ MASTER PLAN ═══════════════════════════════════════════
  function planOf(s) { return db.plans[s] || null; }
  function planAnalysis(s, plan) { return L.analyze(all(), { supplier: s, from: plan && plan.from, to: plan && plan.to }, db.rules); }
  function renderPlan(main) {
    var arg = routeArg();
    var sups = suppliers();
    if (arg && sups.indexOf(arg) >= 0) { ui.plan.supplier = arg; ui.plan.parsed = null; history.replaceState(null, '', '#/plan'); }
    if (!db.records.length) { emptyNotice(main, 'MASTER PLAN'); return; }
    if (sups.indexOf(ui.plan.supplier) < 0) {
      var an0 = L.analyze(all(), {}, db.rules);
      ui.plan.supplier = an0.issues.length ? an0.issues[0].supplier : sups[0];
    }
    var s = ui.plan.supplier, plan = planOf(s);
    var b = L.dateBounds(db.records.filter(function (r) { return r.supplier === s; }));
    main.appendChild(pageHead('MASTER PLAN', '협력사 공정불량을 분석해 불량유형별 개선대책·담당부서·추진 일정 초안을 만들고, 간트차트로 계획과 실적을 관리합니다.'));

    // 조건
    var sup = select(sups, s, { 'aria-label': '협력사' });
    sup.addEventListener('change', function () { ui.plan.supplier = sup.value; ui.plan.parsed = null; ui.plan.aiText = ''; render(); });
    var base = dateInput(plan ? plan.baseDate : nextMonday(), { 'aria-label': '계획 착수 기준일' });
    var from = dateInput(plan ? plan.from : '', { 'aria-label': '분석 시작일' });
    var to = dateInput(plan ? plan.to : '', { 'aria-label': '분석 종료일' });
    if (!plan) { from.value = b.min; to.value = b.max; }
    main.appendChild(h('section', { class: 'card plan-ctl' },
      h('div', { class: 'filters' },
        field('협력사', sup), field('계획 착수 기준일', base),
        field('분석기간 — 시작', from), field('끝', to),
        h('div', { class: 'field' }, h('span', { 'aria-hidden': 'true' }, ' '), h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { generate(s, base.value, from.value, to.value); } }, icon('chip'), '분석 및 초안 생성')))));

    if (!plan) {
      var an1 = L.analyze(all(), { supplier: s }, db.rules);
      main.appendChild(h('section', { class: 'card' }, h('h2', null, s + ' — 아직 Master Plan 이 없습니다'),
        h('p', null, '공정불량 ' + an1.total + '건 · 개선 필요 항목 ' + an1.issues.length + '건' + (an1.issues.length ? ' (' + an1.issues.map(function (i) { return i.type; }).join(', ') + ')' : '') + '.'),
        h('p', { class: 'note' }, '「분석 및 초안 생성」을 누르면 개선 필요 항목의 불량유형마다 설정의 기본 대책으로 초안을 만듭니다(규칙에 걸린 항목이 없으면 건수 상위 3개 유형). 초안은 AI 로 구체화하고 직접 고친 뒤 「계획 확정」합니다.')));
      return;
    }
    var an = planAnalysis(s, plan);
    var stt = L.planStats(plan.rows, today());
    var iss = an.issues.filter(function (i) { return i.supplier === s; });
    main.appendChild(h('section', { class: 'dkpis six' },
      kpi('불량 발생 건수', 'doc', fmt(an.total), '건', an.from + ' ~ ' + an.to),
      kpi('주요 불량 유형 수', 'alert', iss.length, '개', '개선 필요 항목', 'warn'),
      kpi('개선 활동 수', 'list', stt.total, '건'),
      kpi('추진 기간 일수', 'cal', stt.days, '일', stt.start ? stt.start + ' ~ ' + stt.end : ''),
      kpi('완료 / 전체', 'check', stt.done + ' / ' + stt.total, '', '완료율 ' + stt.rate + '%', 'ok'),
      kpi('미완료 / 보류', 'pause', stt.open + ' / ' + stt.hold, '', stt.overdue ? '계획일 지난 미완료 ' + stt.overdue + '건' : '지연 없음', stt.overdue ? 'bad' : 'hold')));

    main.appendChild(aiSummaryCard(s, plan, an, iss));
    main.appendChild(ganttCard(plan.rows, {
      title: '통합 개선대책 Gantt', state: L.planStateText(plan), scaleKey: ui.plan, showSupplier: false,
      onRow: function (r) { openDrawer(s, r.id); },
      actions: [
        h('button', { type: 'button', class: 'btn', onclick: function () { openDrawer(s, null); } }, icon('plus'), '개선활동 추가'),
        h('button', { type: 'button', class: 'btn btn-ok', onclick: function () { confirmPlanUi(s); } }, icon('tick'), '계획 확정'),
        h('button', { type: 'button', class: 'btn', onclick: function () { reloadDraft(s); } }, icon('redo'), 'AI 초안 다시 불러오기'),
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { downloadReport([s]); } }, icon('down'), 'Excel 보고서(.xlsx) 다운로드')
      ]
    }));
    main.appendChild(aiCard(s, plan, an));
  }
  function generate(s, base, from, to) {
    if (!base) { toast('계획 착수 기준일을 골라 주세요.', true); return; }
    if (from && to && to < from) { toast('분석기간의 끝이 시작보다 빠릅니다.', true); return; }
    function go() {
      var old = planOf(s);
      var an = L.analyze(all(), { supplier: s, from: from, to: to }, db.rules);
      var rows = L.buildDraft(an, db.types, { supplier: s, baseDate: base });
      db.plans[s] = { baseDate: base, from: from, to: to, rows: rows, confirmed: old ? old.confirmed : [], aiSummary: '', aiChecks: [], aiAnswer: '' };
      ui.plan.parsed = null; ui.plan.aiText = '';
      save(); render();
      toast(rows.length ? '초안 ' + rows.length + '건을 만들었습니다. 아래 AI 로 구체화하거나 행을 눌러 고치세요.' : '이 기간에는 초안을 만들 불량유형이 없습니다(미분류만 있음).', !rows.length);
    }
    var p = planOf(s);
    if (p && p.rows.length) confirmBox('초안 다시 만들기', '지금 계획 ' + p.rows.length + '건(실적을 적은 행 ' + p.rows.filter(function (r) { return r.actStart || r.status !== 'open'; }).length + '건 포함)을 지우고 새 초안을 만듭니다.\n확정 이력은 남습니다.', '다시 만들기', go, true);
    else go();
  }
  function reloadDraft(s) {
    var p = planOf(s);
    var fromAi = p && p.aiAnswer;
    confirmBox('AI 초안 다시 불러오기', (fromAi ? '마지막으로 반영한 AI 답으로' : '설정의 기본 대책으로') + ' 초안을 다시 만듭니다.\n지금 계획 ' + p.rows.length + '건은 지워집니다(직접 추가한 행도 포함). 확정 이력은 남습니다.', '다시 불러오기', function () {
      if (fromAi) {
        var r = L.parseAiPlan(p.aiAnswer, { types: db.types, aliases: db.aliases, supplier: s, baseDate: p.baseDate });
        p.rows = r.rows; p.aiSummary = r.summary; p.aiChecks = r.checks;
      } else {
        p.rows = L.buildDraft(planAnalysis(s, p), db.types, { supplier: s, baseDate: p.baseDate });
      }
      save(); render(); toast('초안 ' + p.rows.length + '건을 다시 불러왔습니다.');
    }, true);
  }
  function confirmPlanUi(s) {
    var p = planOf(s);
    if (!p.rows.length) { toast('확정할 개선활동이 없습니다.', true); return; }
    var bad = p.rows.filter(function (r) { return !r.planStart || !r.planEnd; }).length;
    var last = L.lastConfirmed(p);
    if (last && !L.planChangedSince(p)) { toast('확정 v' + last.version + ' 이후 계획이 바뀌지 않았습니다.'); return; }
    confirmBox('계획 확정', s + ' Master Plan ' + p.rows.length + '건을 확정 v' + ((p.confirmed || []).length + 1) + '로 남깁니다.' + (bad ? '\n계획일이 비어 있는 활동이 ' + bad + '건 있습니다.' : '') + '\n확정 뒤에도 실적·상태는 계속 적을 수 있고, 계획을 바꾸면 다시 확정하면 됩니다.', '확정', function () {
      var snap = L.confirmPlan(p, nowIso()); save(); render(); toast('확정 v' + snap.version + ' 로 남겼습니다.');
    });
  }
  function aiSummaryCard(s, plan, an, iss) {
    var top = an.byType.filter(function (t) { return t.type !== L.UNCLASSIFIED; })[0];
    var rule = s + ' 공정불량 ' + an.total + '건 가운데 ' + (top ? top.type + '이(가) ' + top.count + '건(' + top.pct + '%)으로 가장 많습니다.' : '분류된 유형이 없습니다.') +
      (iss.length ? ' 개선 필요 항목은 ' + iss.map(function (i) { return i.type + (i.flags.trend ? '(증가 추세)' : ''); }).join(', ') + '입니다.' : ' 규칙에 걸린 개선 필요 항목은 없습니다.');
    var chips = (iss.length ? iss : an.byType.filter(function (t) { return t.type !== L.UNCLASSIFIED; }).slice(0, 3).map(function (t) { return { type: t.type, count: t.count, flags: {} }; }))
      .map(function (i) { return h('span', { class: 'chip' + (i.flags.trend ? ' up' : '') }, i.type, h('small', null, ' ' + i.count + '건' + (i.flags.trend ? ' · 증가' : ''))); });
    return h('section', { class: 'card ai-sum' },
      h('div', { class: 'card-head' }, h('h2', null, icon('chip', 'ai-chip'), 'AI 불량 분석 요약')),
      h('div', { class: 'chips' }, chips),
      h('p', null, plan.aiSummary ? h('span', null, plan.aiSummary, h('small', { class: 'note' }, ' — AI 답(검토 필요)')) : rule),
      plan.aiChecks && plan.aiChecks.length ? h('details', null, h('summary', null, 'AI 가 적은 확인 필요 ' + plan.aiChecks.length + '건'), h('ul', null, plan.aiChecks.map(function (c) { return h('li', null, c); }))) : null);
  }

  // 간트 — 행마다 위 얇은 막대 = 계획, 아래 굵은 막대 = 실적(상태 색). 칸 폭을 고정해 머리 칸과 막대가 맞습니다.
  function ganttCard(rows, o) {
    var scale = o.scaleKey.scale;
    var cols = L.ganttColumns(rows, scale, { today: today() });
    var W = scale === 'month' ? 64 : 46;
    var t = today();
    var todayIdx = -1, todayX = -1;
    cols.forEach(function (c, i) { if (c.start <= t && t <= c.end) { todayIdx = i; todayX = i * W + (L.daysBetween(c.start, t) + 0.5) / (L.daysBetween(c.start, c.end) + 1) * W; } });
    // 좁은 화면: 유형·담당부서는 개선대책 칸 안에 작게 넣고 그 칸을 왼쪽에 붙여 둡니다(막대를 옆으로 밀어 봐도 어느 행인지 보이게)
    var narrow = document.documentElement.clientWidth < 700;
    var fixed = narrow ? [['개선대책', 148], ['상태', 74]]
      : o.showSupplier ? [['협력사', 96], ['불량유형', 104], ['개선대책', 240], ['담당부서', 92], ['실적 상태', 92]] : [['불량유형', 110], ['개선대책', 260], ['담당부서', 96], ['실적 상태', 92]];
    var colgroup = h('colgroup', null, fixed.map(function (f) { return h('col', { style: 'width:' + f[1] + 'px' }); }), cols.map(function () { return h('col', { style: 'width:' + W + 'px' }); }));
    var groups = L.groupSpans(cols);
    var head = h('thead', null,
      h('tr', null, fixed.map(function (f, i) { return h('th', { rowspan: '2', class: 'g-fix' + (i === 0 ? ' g-stick' : '') }, f[0]); }), groups.map(function (g) { return h('th', { colspan: String(g.n), class: 'g-group' }, g.group); })),
      h('tr', null, cols.map(function (c, i) { return h('th', { class: 'g-col' + (i === todayIdx ? ' g-today' : ''), title: c.full }, c.label); })));
    var tb = h('tbody');
    if (!rows.length) tb.appendChild(h('tr', null, h('td', { colspan: String(fixed.length + cols.length), class: 'note' }, '개선활동이 없습니다. 「개선활동 추가」로 넣으세요.')));
    rows.forEach(function (r) {
      var od = L.isOverdue(r, t);
      var lane = h('div', { class: 'g-lane', style: 'width:' + (cols.length * W) + 'px;background-size:' + W + 'px 100%' });
      var sp = L.span(r.planStart, r.planEnd, cols);
      if (sp) lane.appendChild(h('span', { class: 'g-plan st-' + r.status, style: 'left:' + (sp[0] * W + 2) + 'px;width:' + ((sp[1] - sp[0] + 1) * W - 4) + 'px', title: '계획 ' + r.planStart + ' ~ ' + r.planEnd }));
      if (r.actStart) {
        var ae = r.actEnd || (r.status === 'done' ? r.actStart : (t > r.actStart ? t : r.actStart));
        var sa = L.span(r.actStart, ae, cols);
        if (sa) lane.appendChild(h('span', { class: 'g-act st-' + r.status + (r.actEnd ? '' : ' ongoing'), style: 'left:' + (sa[0] * W + 2) + 'px;width:' + ((sa[1] - sa[0] + 1) * W - 4) + 'px', title: '실적 ' + r.actStart + ' ~ ' + (r.actEnd || '진행 중') }));
      }
      if (todayX >= 0) lane.appendChild(h('span', { class: 'g-now', style: 'left:' + todayX + 'px' }));
      var cells = [];
      if (narrow) cells.push(h('td', { class: 'g-fix g-stick g-act-text' }, h('small', { class: 'note' }, (o.showSupplier ? r.supplier + ' · ' : '') + r.type + ' · ' + r.dept), h('span', { class: 'clip-2' }, r.action)),
        h('td', { class: 'g-fix' }, statusPill(r.status, od)));
      else {
      if (o.showSupplier) cells.push(h('td', { class: 'g-fix g-stick' }, r.supplier));
      cells.push(h('td', { class: 'g-fix' }, r.type), h('td', { class: 'g-fix g-act-text' }, h('span', { class: 'clip-2' }, r.action), r.owner ? h('small', { class: 'note' }, r.owner) : null),
        h('td', { class: 'g-fix' }, r.dept), h('td', { class: 'g-fix' }, statusPill(r.status, od)));
      }
      tb.appendChild(h('tr', { class: 'clickable', tabindex: '0', 'aria-label': r.type + ' — ' + r.action + ' (눌러서 계획·실적 입력)',
        onclick: function () { o.onRow(r); }, onkeydown: function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); o.onRow(r); } } },
        cells, h('td', { colspan: String(cols.length), class: 'g-lane-cell' }, lane)));
    });
    var tableW = fixed.reduce(function (a, f) { return a + f[1]; }, 0) + cols.length * W;
    var scaleBtns = h('div', { class: 'seg', role: 'group', 'aria-label': '간트 칸' }, [['week', '주차'], ['month', '월']].map(function (x) {
      return h('button', { type: 'button', 'aria-pressed': o.scaleKey.scale === x[0] ? 'true' : 'false', onclick: function () { o.scaleKey.scale = x[0]; render(); } }, x[1]);
    }));
    return h('section', { class: 'card gantt-card' },
      h('div', { class: 'card-head' }, h('h2', null, o.title), o.state ? h('span', { class: 'badge ' + (o.state.indexOf('확정') === 0 ? (o.state.indexOf('수정') > 0 ? 'warn' : 'ok') : 'muted') }, o.state) : null, scaleBtns),
      h('div', { class: 'g-legend' }, legend(), h('span', { class: 'note' }, '위 얇은 막대 = 계획 · 아래 굵은 막대 = 실적 · 세로선 = 오늘 · 행을 누르면 오른쪽 칸에서 계획·실적을 적습니다')),
      h('div', { class: 'table-wrap gantt-wrap' }, h('table', { class: 'gantt', style: 'width:' + tableW + 'px' }, colgroup, head, tb)),
      o.actions ? h('div', { class: 'btn-row', style: 'margin-top:14px' }, o.actions) : null);
  }

  // AI 로 구체화 — 사실만 담은 프롬프트(협력사·품번은 별칭) → 복사 → 답 붙여 넣기 → 계획 행으로
  function aiCard(s, plan, an) {
    db.aliases = L.updateAliases(db.aliases, db.records);
    var info = L.buildAiPromptInfo({ supplier: s, an: an, rows: plan.rows.filter(function (r) { return r.source === 'template'; }), types: db.types, depts: db.depts, baseDate: plan.baseDate, aliases: db.mask ? db.aliases : null });
    var mask = h('input', { type: 'checkbox', checked: db.mask, onchange: function (e) { db.mask = e.target.checked; save(); render(); } });
    var box = h('textarea', { class: 'prompt-box', readonly: true, 'aria-label': 'AI 프롬프트', rows: '14' }); box.value = info.prompt;
    var ans = h('textarea', { class: 'prompt-box answer', 'aria-label': 'AI 답 붙여 넣기', rows: '10', placeholder: 'AI 답을 여기에 붙여 넣으세요 — 「요약:」 줄, 「| 불량유형 | 개선대책 | 담당부서 | 시작 | 기간 |」 표, 「확인 필요:」 줄을 읽습니다.' });
    ans.value = ui.plan.aiText;
    ans.addEventListener('input', function () { ui.plan.aiText = ans.value; });
    function readAnswer() {
      var r = L.parseAiPlan(ans.value, { types: db.types, aliases: db.aliases, supplier: s, baseDate: plan.baseDate });
      ui.plan.parsed = r; ui.plan.aiText = ans.value; render();
      if (!r.rows.length) toast('표 줄을 찾지 못했습니다. 답이 「| 불량유형 | 개선대책 | … |」 표 형식인지 확인해 주세요.', true);
    }
    var send = window.AIPanel.sendButton('설정한 AI 서버로 보내기', function () { return info.prompt; }, function (text) { ans.value = text; ui.plan.aiText = text; readAnswer(); }, { toast: toast, settingsHref: '#/settings' });
    var preview = null, p = ui.plan.parsed;
    if (p) {
      var known = L.typeNames(db.types);
      preview = h('div', { class: 'ai-preview' },
        h('h3', null, '읽은 결과 — 개선대책 ' + p.rows.length + '건' + (p.skipped ? ' (읽지 못한 표 줄 ' + p.skipped + '건)' : '')),
        p.summary ? h('p', null, h('strong', null, '요약: '), p.summary) : null,
        p.rows.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
          h('thead', null, h('tr', null, ['불량유형', '개선대책', '담당부서', '계획 기간'].map(function (x) { return h('th', null, x); }))),
          h('tbody', null, p.rows.map(function (r) { return h('tr', null, h('td', { class: 'nw' }, r.type, known.indexOf(r.type) < 0 ? h('span', { class: 'badge warn', style: 'margin-left:4px' }, '사전에 없음') : null), h('td', { class: 'wide' }, r.action), h('td', { class: 'nw' }, r.dept), h('td', { class: 'nw' }, r.planStart + ' ~ ' + r.planEnd)); })))) : null,
        p.checks.length ? h('div', null, h('strong', null, '확인 필요'), h('ul', null, p.checks.map(function (c) { return h('li', null, c); }))) : null,
        p.rows.length ? h('div', { class: 'radios' },
          h('label', { class: 'check' }, h('input', { type: 'radio', name: 'aimode', checked: ui.plan.mode === 'replace', onchange: function () { ui.plan.mode = 'replace'; } }), '기본 대책·이전 AI 행을 이 답으로 바꾸기 (직접 추가한 행은 남김)'),
          h('label', { class: 'check' }, h('input', { type: 'radio', name: 'aimode', checked: ui.plan.mode === 'append', onchange: function () { ui.plan.mode = 'append'; } }), '지금 계획에 더하기')) : null,
        p.rows.length ? h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { applyAi(s, p, ans.value); } }, '계획에 반영')) : null);
    }
    return h('section', { class: 'card ai-card', id: 'aiCard' },
      h('div', { class: 'card-head' }, h('h2', null, icon('chip', 'ai-chip'), 'AI 로 개선대책 구체화')),
      h('ol', { class: 'steps' },
        h('li', null, '아래 프롬프트를 복사해 회사가 허용한 AI(사내 AI 또는 ChatGPT 등)에 붙여 넣습니다. 프롬프트에는 이 도구가 계산한 사실만 들어 있습니다.'),
        h('li', null, 'AI 답 전체를 복사해 「AI 답」 칸에 붙여 넣고 「답 읽기」를 누릅니다.'),
        h('li', null, '읽은 결과를 확인하고 「계획에 반영」 → 간트에서 행을 눌러 담당자·일정을 다듬은 뒤 「계획 확정」합니다.')),
      h('label', { class: 'check' }, mask, '협력사 이름과 품번을 별칭(업체1, 품목1)으로 가리기 — 권장. 대응표는 이 브라우저에만 있고, AI 답의 별칭은 원래 이름으로 되돌립니다.'),
      info.leaks.length ? h('p', { class: 'alert warn' }, '가리지 못한 이름이 프롬프트에 남아 있습니다: ' + info.leaks.join(', ') + ' — 보내기 전에 직접 지워 주세요.') : null,
      h('details', { class: 'prompt-details', open: ui.plan.showPrompt, ontoggle: function (e) { ui.plan.showPrompt = e.target.open; } }, h('summary', null, '프롬프트 보기 (' + info.prompt.length + '자)'), box),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(info.prompt); } }, '프롬프트 복사'), send),
      field('AI 답', ans),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn', onclick: readAnswer }, '답 읽기'), ui.plan.parsed ? h('button', { type: 'button', class: 'btn', onclick: function () { ui.plan.parsed = null; ui.plan.aiText = ''; render(); } }, '답 지우기') : null),
      preview);
  }
  function applyAi(s, p, text) {
    var plan = planOf(s);
    if (ui.plan.mode === 'replace') plan.rows = plan.rows.filter(function (r) { return r.source === 'manual'; }).concat(p.rows);
    else plan.rows = plan.rows.concat(p.rows);
    plan.aiSummary = p.summary; plan.aiChecks = p.checks; plan.aiAnswer = text;
    ui.plan.parsed = null; ui.plan.aiText = '';
    save(); render();
    toast('AI 개선대책 ' + p.rows.length + '건을 계획에 넣었습니다. 행을 눌러 담당자와 일정을 확인하세요.');
  }

  // ── 오른쪽 칸: 계획·실적 입력 ─────────────────────────────
  var drawer = document.getElementById('drawer'), drawerBody = document.getElementById('drawerBody'), drawerBack = document.getElementById('drawerBackdrop'), drawerLast = null;
  function closeDrawer(silent) {
    if (drawer.hidden) return;
    drawer.hidden = true; drawerBack.hidden = true; document.body.classList.remove('drawer-open');
    if (!silent && drawerLast && document.body.contains(drawerLast)) drawerLast.focus();
  }
  document.getElementById('drawerClose').addEventListener('click', function () { closeDrawer(); });
  drawerBack.addEventListener('click', function () { closeDrawer(); });
  function openDrawer(s, id) {
    var plan = planOf(s);
    var orig = id ? plan.rows.filter(function (r) { return r.id === id; })[0] : null;
    var r = orig ? JSON.parse(JSON.stringify(orig)) : L.makeRow({ supplier: s, type: (plan.rows[0] || {}).type || '', action: '', dept: '부품품질팀', planStart: plan.baseDate, planEnd: L.addDays(plan.baseDate, 13), source: 'manual' });
    drawerLast = document.activeElement;
    document.getElementById('drawerTitle').textContent = orig ? '개선활동 — 계획·실적 입력' : '개선활동 추가';
    drawerBody.textContent = '';
    var typeIn = h('input', { value: r.type, list: 'dl-types' });
    var act = h('textarea', { rows: '3' }); act.value = r.action;
    var dept = h('input', { value: r.dept, list: 'dl-depts' });
    var owner = h('input', { value: r.owner, placeholder: '예: 홍길동 책임' });
    var ps = dateInput(r.planStart), pe = dateInput(r.planEnd), as = dateInput(r.actStart), ae = dateInput(r.actEnd);
    var note = h('textarea', { rows: '2' }); note.value = r.note;
    var status = r.status;
    var seg = h('div', { class: 'status-seg', role: 'radiogroup', 'aria-label': '실적 상태' });
    function drawSeg() {
      seg.textContent = '';
      ['done', 'open', 'hold'].forEach(function (k) {
        seg.appendChild(h('button', { type: 'button', role: 'radio', 'aria-checked': status === k ? 'true' : 'false', class: 'st-btn st-' + k, onclick: function () {
          status = k;
          if (k === 'done' && !ae.value) { ae.value = today(); if (!as.value) as.value = ps.value && ps.value <= today() ? ps.value : today(); }
          drawSeg();
        } }, L.STATUS[k]));
      });
    }
    drawSeg();
    var errBox = h('div', { 'aria-live': 'polite' });
    function collect() {
      return Object.assign(r, { type: typeIn.value.trim() || '기타', action: act.value.trim(), dept: dept.value.trim() || '부품품질팀', owner: owner.value.trim(),
        planStart: ps.value, planEnd: pe.value, actStart: as.value, actEnd: ae.value, status: status, note: note.value.trim() });
    }
    function doSave() {
      var x = collect(), errs = L.checkPlanRow(x);
      errBox.textContent = '';
      if (errs.length) { errs.forEach(function (e) { errBox.appendChild(h('p', { class: 'alert warn' }, e)); }); return; }
      if (orig) Object.assign(orig, x); else plan.rows.push(x);
      save(); closeDrawer(); render(); toast(orig ? '저장했습니다.' : '개선활동을 더했습니다.');
    }
    add(drawerBody, [
      datalist('dl-types', L.typeNames(db.types)), datalist('dl-depts', db.depts),
      h('p', { class: 'note' }, s + (orig && orig.source !== 'manual' ? ' · ' + (orig.source === 'ai' ? 'AI 답에서 온 행' : '기본 대책에서 온 행') : '')),
      h('h3', null, '계획'),
      field('불량유형', typeIn), field('개선대책', act), h('div', { class: 'form-grid' }, field('담당부서', dept), field('담당자', owner), field('계획 시작', ps), field('계획 완료', pe)),
      h('h3', null, '실적'),
      h('div', { class: 'form-grid' }, field('실적 시작', as), field('실적 완료', ae)),
      h('div', { class: 'field' }, h('span', null, '실적 상태'), seg),
      field('비고(근거·메모)', note),
      errBox,
      h('div', { class: 'btn-row drawer-actions' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: doSave }, '저장'),
        orig ? h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          plan.rows = plan.rows.filter(function (x) { return x.id !== orig.id; }); save(); closeDrawer(); render(); toast('지웠습니다.');
        } }, '이 활동 지우기') : null,
        h('button', { type: 'button', class: 'btn', onclick: function () { closeDrawer(); } }, '닫기'))
    ]);
    drawer.hidden = false; drawerBack.hidden = false; document.body.classList.add('drawer-open');
    drawer.scrollTop = 0;
    setTimeout(function () { (orig ? owner : act).focus(); }, 30);
  }

  // ══ 추진현황 ═══════════════════════════════════════════════
  function renderProgress(main) {
    main.appendChild(pageHead('추진현황', '모든 협력사의 개선활동을 상태별로 봅니다. 행을 누르면 담당자가 실적과 상태를 적을 수 있습니다.'));
    var sups = Object.keys(db.plans).filter(function (s) { return db.plans[s].rows.length; }).sort();
    if (!sups.length) {
      main.appendChild(h('section', { class: 'card' }, h('p', null, '아직 Master Plan 이 없습니다.'), h('a', { class: 'btn btn-primary', href: '#/plan' }, 'MASTER PLAN 으로')));
      return;
    }
    var st = ui.prog, t = today();
    if (st.supplier && sups.indexOf(st.supplier) < 0) st.supplier = '';
    var sup = select([['', '전체 협력사']].concat(sups.map(function (s) { return [s, s]; })), st.supplier, { 'aria-label': '협력사' });
    sup.addEventListener('change', function () { st.supplier = sup.value; render(); });
    var ss = select([['', '전체 상태'], ['done', '완료'], ['open', '미완료'], ['hold', '보류']], st.status, { 'aria-label': '상태' });
    ss.addEventListener('change', function () { st.status = ss.value; render(); });
    var od = h('input', { type: 'checkbox', checked: st.overdue, onchange: function (e) { st.overdue = e.target.checked; render(); } });
    main.appendChild(h('section', { class: 'card' }, h('div', { class: 'filters' }, field('협력사', sup), field('상태', ss), h('label', { class: 'check' }, od, '계획일 지난 미완료만'))));

    var rows = [];
    sups.forEach(function (s) { if (!st.supplier || st.supplier === s) rows = rows.concat(db.plans[s].rows); });
    var ps = L.planStats(rows, t);
    main.appendChild(h('section', { class: 'dkpis six' },
      kpi('전체 개선활동', 'list', ps.total, '건'),
      kpi('완료', 'check', ps.done, '건', '완료율 ' + ps.rate + '%', 'ok'),
      kpi('미완료', 'clock', ps.open, '건', null, 'open'),
      kpi('보류', 'pause', ps.hold, '건', null, 'hold'),
      kpi('지연', 'alert', ps.overdue, '건', '계획 완료일이 지난 미완료', ps.overdue ? 'bad' : ''),
      kpi('추진 기간', 'cal', ps.days, '일', ps.start ? ps.start + ' ~ ' + ps.end : '')));

    var tb = h('tbody');
    sups.forEach(function (s) {
      var p = db.plans[s], x = L.planStats(p.rows, t);
      tb.appendChild(h('tr', null, h('th', { scope: 'row' }, h('a', { href: '#/plan/' + encodeURIComponent(s) }, s)), h('td', null, L.planStateText(p)), h('td', { class: 'stack-cell' }, stackBar(x)),
        h('td', { class: 'num' }, x.done), h('td', { class: 'num' }, x.open), h('td', { class: 'num' }, x.hold), h('td', { class: 'num' }, x.overdue ? h('span', { class: 'late' }, x.overdue) : '0'), h('td', { class: 'nw' }, x.start ? x.start + ' ~ ' + x.end : '')));
    });
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '협력사별'), h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
      h('thead', null, h('tr', null, ['협력사', '계획 상태', '진행', '완료', '미완료', '보류', '지연', '기간'].map(function (x) { return h('th', null, x); }))), tb)), legend()));

    var shown = rows.filter(function (r) { return (!st.status || r.status === st.status) && (!st.overdue || L.isOverdue(r, t)); })
      .sort(function (a, b) { return (a.planEnd || '9') < (b.planEnd || '9') ? -1 : (a.planEnd || '9') > (b.planEnd || '9') ? 1 : 0; });
    main.appendChild(ganttCard(shown, { title: '개선활동 ' + shown.length + '건', scaleKey: st, showSupplier: true, onRow: function (r) { openDrawer(r.supplier, r.id); } }));
  }

  // ══ 보고서 출력 ════════════════════════════════════════════
  function downloadReport(list) {
    var sheets = L.reportSheets({ suppliers: list, records: all(), plans: db.plans, rules: db.rules, today: today() });
    var name = 'MasterPlan_보고서_' + (list.length === 1 ? list[0].replace(/\s+/g, '') : '전체') + tag() + '_' + today() + '.xlsx';
    writeXlsx(name, sheets, { '요약': [22, 40, 10, 10], 'Master Plan': [12, 5, 14, 40, 12, 10, 11, 11, 11, 11, 8, 6].concat(Array(200).fill(9)), '개선대책 상세': [12, 14, 10, 40, 40, 40, 12, 10, 24, 24, 8, 12], '불량 상세': [12, 12, 12, 14, 40, 14, 20], '확정 이력': [12, 10, 18, 12] });
    toast(name + ' 을 내려받았습니다.');
  }
  function renderReport(main) {
    main.appendChild(pageHead('보고서 출력', '확정한 Master Plan 과 개선대책 상세를 보고서 형식의 Excel 로 내보냅니다.'));
    var sups = Object.keys(db.plans).filter(function (s) { return db.plans[s].rows.length; }).sort();
    if (!sups.length) {
      main.appendChild(h('section', { class: 'card' }, h('p', null, '아직 Master Plan 이 없습니다. 보고서는 Master Plan 을 만든 협력사만 나옵니다.'), h('a', { class: 'btn btn-primary', href: '#/plan' }, 'MASTER PLAN 으로')));
      return;
    }
    var st = ui.rep;
    if (st.supplier && sups.indexOf(st.supplier) < 0) st.supplier = '';
    var sel = select([['', '전체 (' + sups.length + '곳)']].concat(sups.map(function (s) { return [s, s]; })), st.supplier, { 'aria-label': '보고서 협력사' });
    sel.addEventListener('change', function () { st.supplier = sel.value; render(); });
    var list = st.supplier ? [st.supplier] : sups;
    var sheets = L.reportSheets({ suppliers: list, records: all(), plans: db.plans, rules: db.rules, today: today() });
    var unconfirmed = list.filter(function (s) { return !L.lastConfirmed(db.plans[s]) || L.planChangedSince(db.plans[s]); });
    main.appendChild(h('section', { class: 'card' },
      h('div', { class: 'filters' }, field('협력사', sel)),
      unconfirmed.length ? h('p', { class: 'alert warn' }, '확정하지 않았거나 확정 뒤 계획을 바꾼 협력사: ' + unconfirmed.join(', ') + ' — 보고서의 「계획 상태」에 그대로 적힙니다.') : h('p', { class: 'alert info' }, '모두 확정한 계획입니다.'),
      h('table', { class: 'list sheets' }, h('thead', null, h('tr', null, h('th', null, '시트'), h('th', null, '내용'), h('th', { class: 'num' }, '행'))),
        h('tbody', null, [['요약', '협력사별 분석 기간·계획 상태·KPI·불량유형 파레토·개선 필요 항목'], ['Master Plan', '개선활동별 계획·실적·상태 + 주차 칸(○ 계획 · ● 실적)'], ['개선대책 상세', '불량유형별 발생 건수·선정 근거·대표 문제점·대책·담당'], ['불량 상세', '분석 기간의 공정불량 Raw data 와 분류 근거'], ['확정 이력', '확정 버전·일시']].map(function (x) {
          return h('tr', null, h('td', { class: 'nw' }, x[0]), h('td', null, x[1]), h('td', { class: 'num' }, Math.max(0, sheets[x[0]].length - 1)));
        }))),
      h('div', { class: 'btn-row', style: 'margin-top:14px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { downloadReport(list); } }, icon('down'), 'Excel 보고서(.xlsx) 다운로드'),
        h('button', { type: 'button', class: 'btn', onclick: function () { window.print(); } }, '인쇄 / PDF'))));
    // 미리보기(인쇄에도 씀)
    var pv = h('section', { class: 'card report-preview' }, h('h2', null, '협력사 공정불량 분석 및 품질개선 Master Plan 보고서'), h('p', { class: 'note' }, '작성일 ' + today()));
    list.forEach(function (s) {
      var p = db.plans[s], an = planAnalysis(s, p), x = L.planStats(p.rows, today());
      pv.appendChild(h('h3', null, s + ' — ' + L.planStateText(p)));
      pv.appendChild(h('p', null, '분석 기간 ' + an.from + ' ~ ' + an.to + ' · 공정불량 ' + an.total + '건 · 개선 필요 항목 ' + an.issues.length + '건 · 개선활동 ' + x.total + '건(완료 ' + x.done + ' · 미완료 ' + x.open + ' · 보류 ' + x.hold + (x.overdue ? ' · 지연 ' + x.overdue : '') + ')'));
      if (p.aiSummary) pv.appendChild(h('p', { class: 'note' }, 'AI 요약(검토 필요): ' + p.aiSummary));
      pv.appendChild(h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
        h('thead', null, h('tr', null, ['No', '불량유형', '개선대책', '담당부서', '담당자', '계획', '실적', '상태'].map(function (c) { return h('th', null, c); }))),
        h('tbody', null, p.rows.map(function (r, i) {
          return h('tr', null, h('td', { class: 'num' }, i + 1), h('td', { class: 'nw' }, r.type), h('td', { class: 'wide' }, r.action), h('td', { class: 'nw' }, r.dept), h('td', null, r.owner),
            h('td', { class: 'nw' }, (r.planStart || '') + ' ~ ' + (r.planEnd || '')), h('td', { class: 'nw' }, r.actStart ? r.actStart + ' ~ ' + (r.actEnd || '') : ''), h('td', null, statusPill(r.status, L.isOverdue(r, today()))));
        })))));
    });
    main.appendChild(pv);
  }

  // ══ 설정 ═══════════════════════════════════════════════════
  function renderSettings(main) {
    main.appendChild(pageHead('설정', '불량유형 키워드 사전·기본 대책, 담당부서, 개선 필요 항목 기준, AI 연결, 데이터 백업'));
    main.appendChild(dictCard());
    var deptBox = h('textarea', { rows: '2' }); deptBox.value = db.depts.join(', ');
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '담당부서 목록'),
      h('p', { class: 'note' }, '쉼표로 나눕니다. 오른쪽 칸의 담당부서 고르기와 AI 프롬프트에 쓰입니다. 실제 부서 목록을 받으면 바꿉니다(가정).'),
      field('담당부서', deptBox),
      h('div', { class: 'btn-row', style: 'margin-top:10px' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
        var d = L.parseKeywords(deptBox.value); if (!d.length) { toast('부서를 하나 이상 적어 주세요.', true); return; }
        db.depts = d; save(); toast('저장했습니다.');
      } }, '저장'))));
    main.appendChild(rulesCard());
    var al = db.aliases || { suppliers: {}, parts: {} };
    main.appendChild(h('section', { class: 'card' }, h('h2', null, 'AI 프롬프트 가리기'),
      h('label', { class: 'check' }, h('input', { type: 'checkbox', checked: db.mask, onchange: function (e) { db.mask = e.target.checked; save(); toast(db.mask ? '협력사·품번을 별칭으로 가립니다.' : '프롬프트에 실제 이름이 들어갑니다.', !db.mask); } }), '협력사 이름과 품번을 별칭으로 가리기(기본·권장)'),
      h('p', { class: 'note' }, '대응표: 협력사 ' + Object.keys(al.suppliers).length + '곳 · 품번 ' + Object.keys(al.parts).length + '개. 같은 이름은 늘 같은 별칭이 됩니다. 대응표는 이 브라우저와 JSON 백업에만 있습니다.')));
    main.appendChild(window.AIPanel.settingsCard({ toast: toast }));
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '사외망 접속과 공유'),
      h('ul', null,
        h('li', null, '이 도구는 GitHub Pages(정적 웹)라 사외망·휴대폰에서도 주소로 열 수 있습니다.'),
        h('li', null, '다만 데이터는 각 브라우저 안에만 저장됩니다. 다른 사람·다른 PC 와 같은 데이터를 보려면 지금은 JSON 백업 파일을 주고받아야 합니다.'),
        h('li', null, '팀·협력사가 함께 보는 공유 DB 와 로그인(협력사별 권한)은 2단계에서 다룹니다. 협력사도 보는지에 따라 설계가 달라져 확인을 요청드렸습니다(기획서 10장).'))));
    main.appendChild(dataCard());
  }
  function dictCard() {
    var counts = {};
    all().forEach(function (r) { counts[r.type] = (counts[r.type] || 0) + 1; });
    var list = h('div', { class: 'dict' });
    var editors = [];
    db.types.forEach(function (t, i) {
      var name = h('input', { value: t.name, 'aria-label': '유형 이름' });
      var kw = h('textarea', { rows: '2', 'aria-label': '키워드' }); kw.value = t.keywords.join(', ');
      var dept = h('input', { value: t.dept, list: 'dl-depts2', 'aria-label': '기본 담당부서' });
      var acts = h('textarea', { rows: String(Math.max(3, t.actions.length + 1)), class: 'mono', 'aria-label': '기본 대책' }); acts.value = L.formatActionLines(t.actions);
      editors.push({ name: name, kw: kw, dept: dept, acts: acts });
      function move(d) { var j = i + d; if (j < 0 || j >= db.types.length) return; readEditors(); var x = db.types[i]; db.types[i] = db.types[j]; db.types[j] = x; save(); render(); }
      list.appendChild(h('details', { class: 'dict-item' },
        h('summary', null, h('span', { class: 'dict-n' }, String(i + 1)), h('strong', null, t.name), h('span', { class: 'note' }, ' · 키워드 ' + t.keywords.length + '개 · 대책 ' + t.actions.length + '개 · 분류 ' + (counts[t.name] || 0) + '건')),
        h('div', { class: 'form-grid' }, field('유형 이름', name), field('기본 담당부서', dept)),
        field('키워드 (쉼표로 나눔)', kw, '띄어쓰기·대소문자는 무시합니다. 여러 유형이 걸리면 가장 긴 키워드가 이기고, 길이가 같으면 위에 있는 유형이 이깁니다.'),
        field('기본 대책 — 한 줄에 「대책 | 담당부서 | 시작(기준일부터 몇 주 뒤) | 기간(주)」', acts),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn btn-sm2', disabled: i === 0, onclick: function () { move(-1); } }, '위로'),
          h('button', { type: 'button', class: 'btn btn-sm2', disabled: i === db.types.length - 1, onclick: function () { move(1); } }, '아래로'),
          h('button', { type: 'button', class: 'btn btn-sm2 btn-danger', onclick: function () {
            confirmBox('유형 지우기', '「' + t.name + '」을 사전에서 지웁니다. 이 유형으로 분류되던 ' + (counts[t.name] || 0) + '건은 다른 유형이나 미분류가 됩니다.', '지우기', function () { readEditors(); db.types.splice(i, 1); save(); render(); }, true);
          } }, '지우기'))));
    });
    function readEditors() {
      var errs = [], names = {};
      editors.forEach(function (e, i) {
        var n = e.name.value.trim();
        if (!n || n === L.UNCLASSIFIED) errs.push((i + 1) + '번 유형 이름을 적어 주세요(「미분류」는 쓸 수 없음).');
        else if (names[n]) errs.push('유형 이름 「' + n + '」이 두 번 있습니다.');
        names[n] = 1;
      });
      if (errs.length) { toast(errs[0], true); return false; }
      var renamed = {};
      editors.forEach(function (e, i) {
        var t = db.types[i], n = e.name.value.trim();
        if (t.name !== n) renamed[t.name] = n;
        t.name = n; t.keywords = L.parseKeywords(e.kw.value); t.dept = e.dept.value.trim() || '협력사'; t.actions = L.parseActionLines(e.acts.value, t.dept);
      });
      // 이름을 바꾸면 직접 고른 유형·계획 행도 따라갑니다
      if (Object.keys(renamed).length) {
        db.records.forEach(function (r) { if (r.typeManual && renamed[r.typeManual]) r.typeManual = renamed[r.typeManual]; });
        Object.keys(db.plans).forEach(function (s) { db.plans[s].rows.forEach(function (r) { if (renamed[r.type]) r.type = renamed[r.type]; }); });
      }
      return true;
    }
    return h('section', { class: 'card' }, h('h2', null, '불량유형 키워드 사전 · 기본 대책'),
      h('p', { class: 'note' }, '문제점 글에 키워드가 들어 있으면 그 불량유형으로 나눕니다. 기본 대책은 Master Plan 초안에 들어갑니다. 지금 값은 1단계 가정입니다 — 실제 불량유형 코드표와 담당부서 목록을 받으면 그대로 옮깁니다(기획서 10장 ②).'),
      datalist('dl-depts2', db.depts),
      list,
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { if (readEditors()) { save(); toast('사전을 저장했습니다. 분류를 다시 계산했습니다.'); render(); } } }, '사전 저장'),
        h('button', { type: 'button', class: 'btn', onclick: function () {
          if (!readEditors()) return;
          var n = '새 유형', k = 1; while (L.typeNames(db.types).indexOf(n) >= 0) n = '새 유형 ' + (++k);
          db.types.push({ name: n, dept: '협력사', keywords: [], actions: [] }); save(); render();
        } }, '유형 더하기'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          confirmBox('기본 사전으로 되돌리기', '불량유형 사전과 기본 대책을 처음 값으로 되돌립니다. 직접 더한 키워드·유형은 사라집니다.', '되돌리기', function () { db.types = L.defaultTypes(); save(); render(); }, true);
        } }, '기본값으로')));
  }
  function rulesCard() {
    var r = db.rules, ins = {};
    function num(k, label, hint) { ins[k] = h('input', { type: 'number', min: '1', value: String(r[k]), inputmode: 'numeric' }); return field(label, ins[k], hint); }
    return h('section', { class: 'card' }, h('h2', null, '개선 필요 항목 기준'),
      h('p', { class: 'note' }, '협력사 × 불량유형마다 셋 중 하나라도 걸리면 개선 필요 항목으로 올립니다. 시작값은 가정이며, 부품품질팀 기준을 받으면 바꿉니다.'),
      h('div', { class: 'form-grid three' },
        num('repeatMin', '반복 발생 — 기간 안 건수 이상'),
        num('partRepeatMin', '반복 품번 — 같은 품번 건수 이상'),
        num('trendMonths', '증가 추세 — 비교 기간(개월)', '최근 N개월과 그 직전 N개월을 비교'),
        num('trendRate', '증가 추세 — 늘어난 비율(%) 이상'),
        num('trendMin', '증가 추세 — 최근 기간 최소 건수')),
      h('div', { class: 'btn-row', style: 'margin-top:10px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          var o = {}; Object.keys(ins).forEach(function (k) { o[k] = ins[k].value; });
          db.rules = L.cleanRules(o); save(); toast('기준을 저장했습니다.'); render();
        } }, '저장'),
        h('button', { type: 'button', class: 'btn', onclick: function () { db.rules = L.cleanRules({}); save(); render(); } }, '시작값으로')));
  }
  function dataCard() {
    var input = h('input', { type: 'file', accept: '.json,application/json', onchange: function (e) {
      var f = e.target.files[0]; e.target.value = '';
      if (!f) return;
      f.text().then(function (t) {
        var o = JSON.parse(t), d = L.restoreDb(o && o.db ? o.db : o);
        confirmBox('백업 되살리기', f.name + ' — Raw data ' + d.records.length + '건, Master Plan ' + Object.keys(d.plans).length + '곳.\n지금 데이터를 이 백업으로 바꿉니다.', '되살리기', function () { db = d; save(); toast('되살렸습니다.'); render(); }, true);
      }).catch(function () { toast('JSON 백업 파일을 읽지 못했습니다.', true); });
    } });
    return h('section', { class: 'card' }, h('h2', null, '데이터 백업'),
      h('p', { class: 'note' }, 'Raw data·사전·Master Plan·확정 이력·별칭 대응표를 JSON 한 파일로 받습니다. AI 키는 넣지 않습니다. 다른 PC 로 옮기거나 다른 사람에게 줄 때 씁니다.'),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          download('data10-07_백업' + tag() + '_' + today() + '.json', new Blob([JSON.stringify({ app: 'data10-07', at: nowIso(), db: db }, null, 1)], { type: 'application/json' }));
        } }, 'JSON 백업 받기'),
        h('label', { class: 'btn file-btn' }, '백업 되살리기', input),
        db.records.length ? h('button', { type: 'button', class: 'btn btn-danger', onclick: clearAll }, '데이터 전부 지우기') : null));
  }

  render();
})();
