let rdData = null;

// 활성 탭/뷰 상태
const state = {
  b:   { tab: null, view: 'total' },
  b620:{ tab: 'B620', view: 'total' },
  rg:  { tab: null, view: 'total' },
};

window.addEventListener('DOMContentLoaded', () => {
  fetch('/api/rawdata-stats', { cache: 'no-store' })
    .then(r => r.json())
    .then(data => {
      if (!data.ok) { showError('데이터 오류'); return; }
      rdData = data;
      document.getElementById('loadingBar').style.display = 'none';
      document.getElementById('pageContent').style.display = 'block';
      initB();
      initB620();
      initRegional();
    })
    .catch(e => showError('서버 오류: ' + e.message));
});

function showError(msg) {
  const el = document.getElementById('loadingBar');
  el.textContent = msg;
  el.style.color = 'var(--primary)';
}

// ── 공통 유틸 ──────────────────────────────────────────
function summaryHtml(tabs_data) {
  return Object.entries(tabs_data).map(([tab, td]) =>
    `<div class="summary-card">
      <div class="s-name">${tab}</div>
      <div class="s-total">${td.total.toLocaleString()}</div>
      <div class="s-sub">건</div>
    </div>`
  ).join('');
}

function tabGroupHtml(tabs, activeTab, prefix, cls) {
  return tabs.map(t =>
    `<button class="tab-btn ${cls}" data-tab="${t}" onclick="${prefix}SetTab('${t}')"
      style="${t===activeTab?'':''}">
      ${t}${t===activeTab?' ✓':''}
    </button>`
  ).join('');
}

function viewToggleHtml(views, activeView, onchangeFn) {
  return views.map(([v, label]) =>
    `<button class="view-btn ${v===activeView?'active':''}" onclick="${onchangeFn}('${v}')">${label}</button>`
  ).join('');
}

/**
 * 유형별 월별 테이블 렌더링
 * data: { months, tabs }
 * tab: 선택된 탭명
 * viewKey: 'total' | '접수오류유형' | '현장처리유형' | '단말기접수유형' | '단말기현장처리'
 */
function renderTable(tableId, data, tab, viewKey) {
  const tbl = document.getElementById(tableId);
  if (!data || !data.tabs[tab]) { tbl.innerHTML = ''; return; }
  const td = data.tabs[tab];
  const months = data.months;

  if (viewKey === 'total') {
    // 월별 총건수 1행
    const cells = months.map(m => `<td class="total-col">${(td.by_month[m]||0).toLocaleString()}</td>`).join('');
    tbl.innerHTML = `
      <thead><tr>
        <th>구분</th>
        ${months.map(m=>`<th>${m.replace('년 ','년<br>')}</th>`).join('')}
        <th>합계</th>
      </tr></thead>
      <tbody>
        <tr>
          <td>전체</td>
          ${cells}
          <td class="total-col">${td.total.toLocaleString()}</td>
        </tr>
      </tbody>`;
    return;
  }

  const typeData = td[viewKey];
  if (!typeData) { tbl.innerHTML = `<tr><td class="empty-msg">데이터 없음</td></tr>`; return; }

  const { types, by_month, total_by_type } = typeData;
  // 합계 기준 내림차순 정렬
  const sorted = [...types].sort((a,b) => (total_by_type[b]||0) - (total_by_type[a]||0));

  const grandTotal = months.reduce((s,m) => s+(td.by_month[m]||0), 0);
  const monthTotals = {};
  months.forEach(m => {
    monthTotals[m] = sorted.reduce((s,t) => s+(by_month[t]?.[m]||0), 0);
  });

  const rows = sorted.map(type => {
    const cells = months.map(m => {
      const v = by_month[type]?.[m] || 0;
      return `<td>${v > 0 ? v.toLocaleString() : '<span style="color:var(--border)">-</span>'}</td>`;
    }).join('');
    return `<tr>
      <td title="${type}">${type.length > 22 ? type.slice(0,22)+'…' : type}</td>
      ${cells}
      <td class="total-col">${(total_by_type[type]||0).toLocaleString()}</td>
    </tr>`;
  }).join('');

  const footCells = months.map(m => `<td>${(monthTotals[m]||0).toLocaleString()}</td>`).join('');

  tbl.innerHTML = `
    <thead><tr>
      <th>${viewKey}</th>
      ${months.map(m=>`<th>${m.replace('년 ','년<br>')}</th>`).join('')}
      <th>합계</th>
    </tr></thead>
    <tbody>
      ${rows}
      <tr>
        <td>합계</td>
        ${footCells}
        <td>${grandTotal.toLocaleString()}</td>
      </tr>
    </tbody>`;
}

// ── B800/B700/B710 ────────────────────────────────────
function initB() {
  const data = rdData.b_series;
  if (!data) { document.getElementById('bSeriesSection').style.display='none'; return; }

  const tabs = ['B700','B710','B800'].filter(t => data.tabs[t]);
  if (!tabs.length) return;
  state.b.tab = tabs[0];

  document.getElementById('bSummary').innerHTML = summaryHtml(
    Object.fromEntries(tabs.map(t => [t, data.tabs[t]]))
  );

  renderBTabGroup();
  renderBViewToggle();
  renderBTable();
}

function renderBTabGroup() {
  const data = rdData.b_series;
  const tabs = ['B700','B710','B800'].filter(t => data.tabs[t]);
  document.getElementById('bTabGroup').innerHTML =
    tabs.map(t =>
      `<button class="tab-btn b-tab ${t===state.b.tab?'active':''}" onclick="bSetTab('${t}')">${t}</button>`
    ).join('');
}

function bSetTab(tab) {
  state.b.tab = tab;
  renderBTabGroup();
  renderBTable();
}

function renderBViewToggle() {
  const views = [
    ['total','전체 건수'],
    ['접수오류유형','접수오류유형'],
    ['현장처리유형','현장처리유형'],
  ];
  document.getElementById('bViewToggle').innerHTML =
    views.map(([v,lbl]) =>
      `<button class="view-btn ${v===state.b.view?'active':''}" onclick="bSetView('${v}')">${lbl}</button>`
    ).join('');
}

function bSetView(view) {
  state.b.view = view;
  renderBViewToggle();
  renderBTable();
}

function renderBTable() {
  renderTable('bTable', rdData.b_series, state.b.tab, state.b.view);
}

// ── 공항 B620 ─────────────────────────────────────────
function initB620() {
  const data = rdData.b620;
  if (!data) { document.getElementById('b620Section').style.display='none'; return; }

  const tabs = Object.keys(data.tabs);
  if (!tabs.length) return;
  state.b620.tab = tabs[0];

  document.getElementById('b620Summary').innerHTML = summaryHtml(data.tabs);
  renderB620ViewToggle();
  renderB620Table();
}

function renderB620ViewToggle() {
  const views = [
    ['total','전체 건수'],
    ['접수오류유형','접수오류유형'],
    ['현장처리유형','현장처리유형'],
  ];
  document.getElementById('b620ViewToggle').innerHTML =
    views.map(([v,lbl]) =>
      `<button class="view-btn ${v===state.b620.view?'active':''}" onclick="b620SetView('${v}')">${lbl}</button>`
    ).join('');
}

function b620SetView(view) {
  state.b620.view = view;
  renderB620ViewToggle();
  renderB620Table();
}

function renderB620Table() {
  renderTable('b620Table', rdData.b620, state.b620.tab, state.b620.view);
}

// ── 지역버스 ─────────────────────────────────────────
const RG_ORDER = ['B650','B400','B600','B500','B810','B520D'];

function initRegional() {
  const data = rdData.regional;
  if (!data) { document.getElementById('regionalSection').style.display='none'; return; }

  const tabs = RG_ORDER.filter(t => data.tabs[t]);
  if (!tabs.length) return;
  state.rg.tab = tabs[0];

  document.getElementById('rgSummary').innerHTML = summaryHtml(
    Object.fromEntries(tabs.map(t => [t, data.tabs[t]]))
  );

  renderRgTabGroup();
  renderRgViewToggle();
  renderRgTable();
}

function renderRgTabGroup() {
  const data = rdData.regional;
  const tabs = RG_ORDER.filter(t => data.tabs[t]);
  document.getElementById('rgTabGroup').innerHTML =
    tabs.map(t =>
      `<button class="tab-btn rg-tab ${t===state.rg.tab?'active':''}" onclick="rgSetTab('${t}')">${t}</button>`
    ).join('');
}

function rgSetTab(tab) {
  state.rg.tab = tab;
  renderRgTabGroup();
  renderRgTable();
}

function renderRgViewToggle() {
  const views = [
    ['total','전체 건수'],
    ['단말기접수유형','단말기접수유형'],
    ['단말기현장처리','단말기현장처리'],
  ];
  document.getElementById('rgViewToggle').innerHTML =
    views.map(([v,lbl]) =>
      `<button class="view-btn ${v===state.rg.view?'active':''}" onclick="rgSetView('${v}')">${lbl}</button>`
    ).join('');
}

function rgSetView(view) {
  state.rg.view = view;
  renderRgViewToggle();
  renderRgTable();
}

function renderRgTable() {
  renderTable('rgTable', rdData.regional, state.rg.tab, state.rg.view);
}
