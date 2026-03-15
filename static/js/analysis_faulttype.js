const FT_TAB_META = {
  b_series: { cls: 'b-tab',  views: ['접수오류유형','현장처리유형'] },
  airport:  { cls: 'ap-tab', views: ['접수오류유형','현장처리유형'] },
  regional: { cls: 'rg-tab', views: ['단말기접수유형','단말기현장처리'] },
};

let ftAllData = null;
let ftState = { groupKey: null, tab: null, view: null, year: 'all' };

window.addEventListener('DOMContentLoaded', () => loadFaultTypes());
document.addEventListener('visibilitychange', () => { if (!document.hidden) loadFaultTypes(); });

function loadFaultTypes() {
  fetch('/api/fault-type-stats', { cache:'no-store' })
    .then(r => r.json())
    .then(data => {
      if (!data.ok) {
        document.getElementById('loadingBar').textContent = '데이터 오류: ' + (data.error || '알 수 없는 오류');
        document.getElementById('loadingBar').style.color = 'var(--primary)';
        return;
      }
      ftAllData = data;
      const flat = ftFlatTabs();
      if (!flat.length) {
        document.getElementById('loadingBar').textContent = '데이터가 없습니다.';
        return;
      }
      ftState.groupKey = flat[0].groupKey;
      ftState.tab = flat[0].tab;
      ftState.view = FT_TAB_META[flat[0].groupKey].views[0];
      ftState.year = 'all';

      document.getElementById('loadingBar').style.display = 'none';
      document.getElementById('dashContent').style.display = 'block';

      renderFtTabGroup();
      renderFtViewToggle();
      renderFtYearSeg();
      renderFtTable();
    })
    .catch(e => {
      document.getElementById('loadingBar').textContent = '서버 연결 오류: ' + e.message;
      document.getElementById('loadingBar').style.color = 'var(--primary)';
    });
}

function ftFlatTabs() {
  if (!ftAllData) return [];
  const list = [];
  for (const [gk, meta] of Object.entries(FT_TAB_META)) {
    const gd = ftAllData[gk];
    if (!gd) continue;
    for (const tab of Object.keys(gd.tabs||{})) {
      list.push({ tab, groupKey: gk, cls: meta.cls });
    }
  }
  return list;
}

function renderFtTabGroup() {
  const flat = ftFlatTabs();
  const el = document.getElementById('ftTabGroup');
  el.innerHTML = flat.map(({ tab, groupKey, cls }) =>
    `<button class="tab-btn ${cls} ${tab===ftState.tab?'active':''}" data-tab="${tab}" data-gk="${groupKey}">${tab}</button>`
  ).join('');
  el.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', () => ftSetTab(btn.dataset.tab, btn.dataset.gk));
  });
}

function ftSetTab(tab, groupKey) {
  ftState.tab = tab;
  ftState.groupKey = groupKey;
  ftState.view = FT_TAB_META[groupKey].views[0];
  renderFtTabGroup();
  renderFtViewToggle();
  renderFtTable();
}

function renderFtViewToggle() {
  const views = FT_TAB_META[ftState.groupKey]?.views || [];
  document.getElementById('ftViewToggle').innerHTML =
    `<div class="ft-seg">${views.map(v =>
      `<button class="ft-seg-btn ${v===ftState.view?'active':''}" onclick="ftSetView('${v}')">${v}</button>`
    ).join('')}</div>`;
}

function ftSetView(view) {
  ftState.view = view;
  renderFtViewToggle();
  renderFtTable();
}

// ─── 연도 필터 ────────────────────────────────────────
function renderFtYearSeg() {
  const gd = ftAllData?.[ftState.groupKey];
  if (!gd) return;
  const months = gd.months || [];
  const years = [...new Set(months.map(m => {
    // "2025년 7월" 또는 "25년 7월" 형식 파싱
    const match = m.match(/(\d{2,4})년/);
    if (!match) return null;
    const y = match[1];
    return y.length === 2 ? '20' + y : y;
  }).filter(Boolean))].sort();

  const btns = [{ val: 'all', label: '전체' }, ...years.map(y => ({ val: y, label: y + '년' }))];
  document.getElementById('ftYearSeg').innerHTML =
    `<div class="ft-seg">${btns.map(b =>
      `<button class="ft-seg-btn ${b.val===ftState.year?'active':''}" onclick="ftSetYear('${b.val}')">${b.label}</button>`
    ).join('')}</div>`;
}

function ftSetYear(year) {
  ftState.year = year;
  renderFtYearSeg();
  renderFtTable();
}

// ─── 테이블 렌더링 ────────────────────────────────────
function renderFtTable() {
  const tbl = document.getElementById('ftTable');
  const gd = ftAllData?.[ftState.groupKey];
  if (!gd || !gd.tabs[ftState.tab]) { tbl.innerHTML = ''; return; }

  const td = gd.tabs[ftState.tab];
  let months = gd.months;

  // 연도 필터 적용
  if (ftState.year !== 'all') {
    months = months.filter(m => {
      const match = m.match(/(\d{2,4})년/);
      if (!match) return false;
      const y = match[1];
      const fullY = y.length === 2 ? '20' + y : y;
      return fullY === ftState.year;
    });
  }

  const typeData = td[ftState.view];
  if (!typeData || !typeData.types?.length) {
    tbl.innerHTML = `<tr><td class="empty-msg" colspan="99">데이터 없음</td></tr>`; return;
  }

  const { types, by_month, total_by_type } = typeData;

  // 필터된 월 기준 합계 재계산
  const filteredTotalByType = {};
  types.forEach(type => {
    filteredTotalByType[type] = months.reduce((s, m) => s + (by_month[type]?.[m] || 0), 0);
  });
  const typeTotal = Object.values(filteredTotalByType).reduce((s, v) => s + v, 0);
  const monthTotals = {};
  months.forEach(m => {
    monthTotals[m] = types.reduce((s, t) => s + (by_month[t]?.[m] || 0), 0);
  });

  const rows = types.map(type => {
    const tot = filteredTotalByType[type] || 0;
    if (tot === 0) return '';
    const cells = months.map(m => {
      const v = by_month[type]?.[m] || 0;
      return `<td>${v > 0 ? v.toLocaleString()+'건' : '<span style="color:var(--border)">-</span>'}</td>`;
    }).join('');
    const pct = typeTotal > 0 ? (tot / typeTotal * 100).toFixed(1) : 0;
    return `<tr>
      <td title="${type}">${type.length > 22 ? type.slice(0,22)+'…' : type}</td>
      ${cells}
      <td class="total-col">${tot.toLocaleString()}건<br><span style="font-size:11px;font-weight:600;color:var(--gray)">${pct}%</span></td>
    </tr>`;
  }).join('');

  const footCells = months.map(m => `<td>${(monthTotals[m]||0).toLocaleString()}건</td>`).join('');

  tbl.innerHTML = `
    <thead><tr>
      <th>${ftState.view}</th>
      ${months.map(m=>`<th>${m.replace('년 ','년<br>')}</th>`).join('')}
      <th>합계</th>
    </tr></thead>
    <tbody>
      ${rows}
      <tr><td>합계</td>${footCells}<td>${typeTotal.toLocaleString()}건</td></tr>
    </tbody>`;
}
