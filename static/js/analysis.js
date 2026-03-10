const TAB_ORDER = [
  "서울 B800","서울 B700","서울 B710","공항 B620",
  "대전 B650","세종 B500","제주 B400","포항 B800",
  "상주,영주,예천 B400","안동 B520D","김해 B600"
];
const TAB_COLORS = {
  "서울 B800":"#C8175C","서울 B700":"#E0507A","서울 B710":"#F08098","공항 B620":"#1A4F7A",
  "대전 B650":"#1E7A3C","세종 B500":"#2E9A55","제주 B400":"#3EBA6E","포항 B800":"#4E8E33",
  "상주,영주,예천 B400":"#6EAA55","안동 B520D":"#8ECA77","김해 B600":"#AABB55"
};
function getTabClass(t){return t.startsWith("서울")?"seoul":t.startsWith("공항")?"airport":"regional";}

let allData = null;
let activeCharts = {};
let activeTabs = new Set(TAB_ORDER);

// ─── 초기 로드 ──────────────────────────────────────
function loadDashboard(isReload) {
  fetch('/api/analysis-data', { cache: 'no-store' })
    .then(r => {
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      return r.json();
    })
    .then(data => {
      if (!data.ok) { showError('데이터 오류: ' + (data.error || '알 수 없는 오류')); return; }
      allData = data;
      document.getElementById('loadingBar').style.display = 'none';
      document.getElementById('dashContent').style.display = 'block';

      try { initTabChecks(); } catch(e) { console.error('탭 초기화 오류:', e); }
      try { renderStatsTable(data.all_stats || {}); } catch(e) { console.error('테이블 오류:', e); }
      try { renderDownloads(data.all_stats || {}); } catch(e) { console.error('다운로드 오류:', e); }
      try { initCompareSelects(data.labels || []); } catch(e) { console.error('셀렉트 오류:', e); }
      try { renderCompare(); } catch(e) { console.error('비교 오류:', e); }
      try { loadConfirmedWeeksStats(); } catch(e) { console.error('확정 주차 오류:', e); }
      try { loadFaultTypeSection(); } catch(e) { console.error('장애유형 섹션 오류:', e); }

      if (isReload) {
        // 차트 재생성 (기존 차트 파괴 후 재생성)
        try {
          if (activeCharts.count) { activeCharts.count.destroy(); activeCharts.count = null; }
          if (activeCharts.rate)  { activeCharts.rate.destroy();  activeCharts.rate  = null; }
          renderCharts();
        } catch(e) { console.error('차트 재생성 오류:', e); }
      } else {
        try { renderCharts(); } catch(e) {
          console.error('차트 오류:', e);
          document.getElementById('countChart').parentElement.innerHTML =
            '<p style="color:var(--gray);font-size:12px;padding:20px">차트를 불러올 수 없습니다.</p>';
          document.getElementById('rateChart').parentElement.innerHTML =
            '<p style="color:var(--gray);font-size:12px;padding:20px">차트를 불러올 수 없습니다.</p>';
        }
      }
    })
    .catch(e => { showError('서버 연결 오류: ' + e.message); });
}

window.addEventListener('DOMContentLoaded', () => loadDashboard(false));

// 다른 탭에서 삽입/확정 후 돌아올 때 자동 갱신
document.addEventListener('visibilitychange', () => {
  if (!document.hidden && allData) loadDashboard(true);
});

function showError(msg) {
  document.getElementById('loadingBar').textContent = msg;
  document.getElementById('loadingBar').style.color = 'var(--primary)';
}

// ─── 탭 필터 ─────────────────────────────────────────
function initTabChecks() {
  activeTabs = new Set(TAB_ORDER);
  const wrap = document.getElementById('tabChecks');
  wrap.innerHTML = TAB_ORDER.map(tab => {
    const cls = getTabClass(tab);
    return `<span class="tab-check active ${cls}" data-tab="${tab}">${tab}</span>`;
  }).join('');
  wrap.querySelectorAll('.tab-check').forEach(el => {
    el.addEventListener('click', () => toggleTab(el, el.dataset.tab));
  });
}
function toggleTab(el, tab) {
  if (activeTabs.has(tab)) { activeTabs.delete(tab); el.classList.remove('active'); }
  else { activeTabs.add(tab); el.classList.add('active'); }
  try { updateCharts(); } catch(e) {}
}
function selectAllTabs() {
  activeTabs = new Set(TAB_ORDER);
  document.querySelectorAll('.tab-check').forEach(el => el.classList.add('active'));
  try { updateCharts(); } catch(e) {}
}
function clearAllTabs() {
  activeTabs.clear();
  document.querySelectorAll('.tab-check').forEach(el => el.classList.remove('active'));
  try { updateCharts(); } catch(e) {}
}

// ─── 차트 ────────────────────────────────────────────
function buildDatasets(field) {
  return TAB_ORDER.map(tab => {
    const ds = allData.chart_datasets[tab] || {};
    const vals = (ds[field] || []).map(v => v == null ? 0 : Number(v));
    return {
      label: tab, data: vals,
      borderColor: TAB_COLORS[tab] || '#888',
      backgroundColor: (TAB_COLORS[tab] || '#888') + '33',
      borderWidth: 2, pointRadius: 3, tension: 0.3, fill: false,
      hidden: !activeTabs.has(tab),
    };
  });
}

function renderCharts() {
  if (typeof Chart === 'undefined') throw new Error('Chart.js not loaded');

  const legendClickHandler = (e, legendItem, legend) => {
    const tab = legendItem.text;
    const isHidden = legend.chart.data.datasets[legendItem.datasetIndex].hidden;
    if (isHidden) activeTabs.add(tab); else activeTabs.delete(tab);
    const el = [...document.querySelectorAll('.tab-check')].find(el => el.dataset.tab === tab);
    if (el) { isHidden ? el.classList.add('active') : el.classList.remove('active'); }
    updateCharts();
  };

  const commonOpts = {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'bottom',
        labels: { font: { size: 11 }, boxWidth: 12, padding: 8 },
        onClick: legendClickHandler,
      },
      tooltip: { mode: 'index', intersect: false },
    },
    interaction: { mode: 'nearest', axis: 'x', intersect: false },
  };

  activeCharts.count = new Chart(document.getElementById('countChart'), {
    type: 'line',
    data: { labels: allData.labels, datasets: buildDatasets('counts') },
    options: {
      ...commonOpts,
      scales: {
        x: { ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });

  activeCharts.rate = new Chart(document.getElementById('rateChart'), {
    type: 'line',
    data: {
      labels: allData.labels,
      datasets: buildDatasets('rates'),
    },
    options: {
      ...commonOpts,
      scales: {
        x: { ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { callback: v => v + '%', font: { size: 10 } } },
      },
    },
  });
}

function updateCharts() {
  if (!allData || !activeCharts.count) return;
  [activeCharts.count, activeCharts.rate].forEach(chart => {
    chart.data.datasets.forEach(ds => {
      ds.hidden = !activeTabs.has(ds.label);
    });
    chart.update();
  });
}

// ─── 월 비교 셀렉트 ──────────────────────────────────
function initCompareSelects(labels) {
  const opts = labels.map((l, i) => `<option value="${i}">${l}</option>`).join('');
  document.getElementById('monthSelA').innerHTML = opts;
  document.getElementById('monthSelB').innerHTML = opts;
  if (labels.length >= 2) {
    document.getElementById('monthSelA').value = String(labels.length - 2);
    document.getElementById('monthSelB').value = String(labels.length - 1);
  }
}


function renderCompare() {
  if (!allData) return;
  const sortedKeys = Object.keys(allData.all_stats).sort();
  const idxA = parseInt(document.getElementById('monthSelA').value || '0');
  const idxB = parseInt(document.getElementById('monthSelB').value || '1');
  const statA = allData.all_stats[sortedKeys[idxA]];
  const statB = allData.all_stats[sortedKeys[idxB]];
  if (!statA || !statB) return;

  // 최대 건수 → 변화량 바 길이 계산용
  const maxCount = Math.max(
    ...TAB_ORDER.map(t => Math.max(
      (statA.tabs[t] || {}).count || 0,
      (statB.tabs[t] || {}).count || 0
    )), 1
  );

  function cellA(tab) {
    const d = (statA.tabs || {})[tab] || {};
    const cnt = d.count || 0;
    const rate = d.fault_rate;
    const rateCls = rate == null ? '' : rate >= 1 ? 'bad' : 'ok';
    const rateTxt = rate != null ? `${rate}%` : '-';
    return `<td>
      <span class="cmp-count">${cnt}건</span>
      <span class="cmp-rate ${rateCls}">장애율 ${rateTxt}</span>
    </td>`;
  }

  function cellB(tab) {
    const d = (statB.tabs || {})[tab] || {};
    const cnt = d.count || 0;
    const rate = d.fault_rate;
    const rateCls = rate == null ? '' : rate >= 1 ? 'bad' : 'ok';
    const rateTxt = rate != null ? `${rate}%` : '-';
    return `<td>
      <span class="cmp-count">${cnt}건</span>
      <span class="cmp-rate ${rateCls}">장애율 ${rateTxt}</span>
    </td>`;
  }

  function cellDiff(tab) {
    const cA = (statA.tabs[tab] || {}).count || 0;
    const cB = (statB.tabs[tab] || {}).count || 0;
    const diff = cB - cA;
    const cls = diff > 0 ? 'diff-up' : diff < 0 ? 'diff-down' : 'diff-same';
    const arrow = diff > 0 ? '▲' : diff < 0 ? '▼' : '→';
    const barW = Math.round(Math.abs(diff) / maxCount * 60);
    const barColor = diff > 0 ? 'var(--primary)' : diff < 0 ? 'var(--green)' : 'var(--gray)';
    const barHtml = diff !== 0 ? `<span class="cmp-bar" style="width:${barW}px;background:${barColor}"></span>` : '';
    const pct = cA > 0 ? ` (${diff > 0 ? '+' : ''}${Math.round(diff / cA * 100)}%)` : '';
    return `<td>
      <span class="cmp-diff ${cls}">${arrow} ${Math.abs(diff)}건${pct}</span>${barHtml}
    </td>`;
  }

  const tA = Object.values(statA.tabs).reduce((s, d) => s + (d.count || 0), 0);
  const tB = Object.values(statB.tabs).reduce((s, d) => s + (d.count || 0), 0);
  const tDiff = tB - tA;
  const tCls = tDiff > 0 ? 'diff-up' : tDiff < 0 ? 'diff-down' : 'diff-same';
  const tArrow = tDiff > 0 ? '▲' : tDiff < 0 ? '▼' : '→';
  const tPct = tA > 0 ? ` (${tDiff > 0 ? '+' : ''}${Math.round(tDiff / tA * 100)}%)` : '';

  const rows = TAB_ORDER.map(tab => {
    const cA = (statA.tabs[tab] || {}).count || 0;
    const cB = (statB.tabs[tab] || {}).count || 0;
    if (cA === 0 && cB === 0) return '';
    return `<tr>
      <td><span class="tab-label-analysis ${getTabClass(tab)}">${tab}</span></td>
      ${cellA(tab)}${cellB(tab)}${cellDiff(tab)}
    </tr>`;
  }).join('');

  document.getElementById('compareGrid').innerHTML = `
    <thead>
      <tr>
        <th>단말기</th>
        <th>${statA.label}</th>
        <th>${statB.label}</th>
        <th>변화량</th>
      </tr>
    </thead>
    <tbody>
      ${rows}
      <tr>
        <td>합계</td>
        <td><span class="cmp-count">${tA}건</span></td>
        <td><span class="cmp-count">${tB}건</span></td>
        <td><span class="cmp-diff ${tCls}">${tArrow} ${Math.abs(tDiff)}건${tPct}</span></td>
      </tr>
    </tbody>`;
}


// ─── 전체 통계 테이블 ────────────────────────────────
function renderStatsTable(all_stats) {
  const sortedKeys = Object.keys(all_stats).sort();
  const monthLabels = sortedKeys.map(k => all_stats[k].label);
  const head = document.getElementById('statsHead');
  const body = document.getElementById('statsBody');

  head.innerHTML =
    '<tr><th>단말기</th>' + monthLabels.map(l => `<th colspan="2">${l}</th>`).join('') + '</tr>' +
    '<tr><th></th>' + monthLabels.map(() => '<th>건수</th><th>장애율</th>').join('') + '</tr>';

  body.innerHTML = TAB_ORDER.map(tab => {
    const cells = sortedKeys.map(k => {
      const d = ((all_stats[k].tabs || {})[tab]) || {};
      const cnt = d.count != null ? d.count : '-';
      const rate = d.fault_rate != null ? d.fault_rate + '%' : '-';
      const rateCls = d.fault_rate != null ? (d.fault_rate >= 1 ? 'rate-bad' : 'rate-ok') : '';
      return `<td>${cnt}</td><td class="rate-cell ${rateCls}">${rate}</td>`;
    }).join('');
    return `<tr><td style="text-align:center"><span class="tab-label-analysis ${getTabClass(tab)}">${tab}</span></td>${cells}</tr>`;
  }).join('');
}

// ─── 다운로드 ────────────────────────────────────────
function renderDownloads(all_stats) {
  const row = document.getElementById('downloadRow');
  const sortedKeys = Object.keys(all_stats).sort();
  // 확정(confirmed) 파일만 표시
  const confirmedKeys = sortedKeys.filter(k => all_stats[k].confirmed);
  row.innerHTML = confirmedKeys.map(k => {
    const fn = all_stats[k].filename;
    const label = all_stats[k].label;
    return `<button class="btn-dl confirmed" onclick="downloadFile('${encodeURIComponent(fn)}')">
      &#128274; ${label} &#11015;
    </button>`;
  }).join('') || '<p style="color:var(--gray);font-size:12px">다운로드 가능한 파일이 없습니다.</p>';
}

function downloadFile(encoded) {
  window.location.href = '/api/download-monthly/' + encoded;
}

// ─── 확정된 주차 현황 ─────────────────────────────────
let confirmedWeeksData = [];

function loadConfirmedWeeksStats() {
  fetch('/api/confirmed-weeks-stats', { cache: 'no-store' })
    .then(r => r.json())
    .then(data => {
      if (!data.ok || !data.weeks || data.weeks.length === 0) return;
      confirmedWeeksData = data.weeks;

      // 드롭다운 채우기 (최신순) — 주차 레이블만 표시, 확정 주차는 🔒
      const sel = document.getElementById('confirmedWeekSel');
      sel.innerHTML = [...data.weeks].reverse().map((w, i) => {
        const lock = w.confirmed ? ' 🔒' : '';
        return `<option value="${data.weeks.length - 1 - i}">${w.week_label}${lock}</option>`;
      }).join('');

      document.getElementById('confirmedWeeksSection').style.display = '';
      renderConfirmedWeekDetail();

      // 주차 비교 드롭다운 (최신순)
      if (data.weeks.length >= 2) {
        const optHtml = [...data.weeks].reverse().map((w, i) => {
          const lock = w.confirmed ? ' 🔒' : '';
          return `<option value="${data.weeks.length - 1 - i}">${w.week_label}${lock}</option>`;
        }).join('');
        document.getElementById('weekSelA').innerHTML = optHtml;
        document.getElementById('weekSelB').innerHTML = optHtml;
        // 기본: 최신 vs 직전
        document.getElementById('weekSelA').value = data.weeks.length - 2;
        document.getElementById('weekSelB').value = data.weeks.length - 1;
        document.getElementById('weekCompareSection').style.display = '';
        renderWeekCompare();
      }
    })
    .catch(e => console.error('confirmed-weeks-stats 오류:', e));
}

function renderConfirmedWeekDetail() {
  if (!confirmedWeeksData.length) return;
  const idx = parseInt(document.getElementById('confirmedWeekSel').value || '0');
  const week = confirmedWeeksData[idx];
  if (!week) return;

  const tbody = document.getElementById('confirmedWeekBody');
  tbody.innerHTML = TAB_ORDER.map(tab => {
    const cnt = week.tab_counts[tab] || 0;
    if (cnt === 0) return '';
    const cls = getTabClass(tab);
    const top3 = week.tab_top3?.[tab] || '-';
    const rate = week.tab_rates?.[tab];
    const rateHtml = rate != null
      ? `<span class="rate-badge ${rate >= 1 ? 'bad' : 'ok'}">${rate}%</span>`
      : '';
    return `<tr>
      <td><span class="tab-label ${cls}">${tab}</span></td>
      <td class="count-cell">${cnt}건 ${rateHtml}</td>
      <td class="period-cell"><span class="week-badge">${week.week_label} <b>${cnt}</b>건</span></td>
      <td class="top3-cell">${top3}</td>
    </tr>`;
  }).join('');
}

// ─── 주차 비교 ────────────────────────────────────────
function renderWeekCompare() {
  if (confirmedWeeksData.length < 2) return;
  const idxA = parseInt(document.getElementById('weekSelA').value);
  const idxB = parseInt(document.getElementById('weekSelB').value);
  const wA = confirmedWeeksData[idxA];
  const wB = confirmedWeeksData[idxB];
  if (!wA || !wB) return;

  const maxCnt = Math.max(
    ...TAB_ORDER.map(t => Math.max(wA.tab_counts[t]||0, wB.tab_counts[t]||0)), 1
  );

  const rows = TAB_ORDER.map(tab => {
    const cA = wA.tab_counts[tab]||0, cB = wB.tab_counts[tab]||0;
    if (cA===0 && cB===0) return '';
    const rA = wA.tab_rates?.[tab], rB = wB.tab_rates?.[tab];
    const diff = cB - cA;
    const cls = diff>0?'diff-up':diff<0?'diff-down':'diff-same';
    const arrow = diff>0?'▲':diff<0?'▼':'→';
    const pct = cA>0?` (${diff>0?'+':''}${Math.round(diff/cA*100)}%)`:'';
    const barW = Math.round(Math.abs(diff)/maxCnt*60);
    const barColor = diff>0?'var(--primary)':diff<0?'var(--green)':'var(--gray)';
    const barHtml = diff!==0?`<span class="cmp-bar" style="width:${barW}px;background:${barColor}"></span>`:'';
    const rateCellA = rA!=null?`<span class="cmp-rate ${rA>=1?'bad':'ok'}">장애율 ${rA}%</span>`:`<span class="cmp-rate">-</span>`;
    const rateCellB = rB!=null?`<span class="cmp-rate ${rB>=1?'bad':'ok'}">장애율 ${rB}%</span>`:`<span class="cmp-rate">-</span>`;
    return `<tr>
      <td><span class="tab-label-analysis ${getTabClass(tab)}">${tab}</span></td>
      <td><span class="cmp-count">${cA}건</span>${rateCellA}</td>
      <td><span class="cmp-count">${cB}건</span>${rateCellB}</td>
      <td><span class="cmp-diff ${cls}">${arrow} ${Math.abs(diff)}건${pct}</span>${barHtml}</td>
    </tr>`;
  }).join('');

  const tA=wA.total||0, tB=wB.total||0, tD=tB-tA;
  const tCls=tD>0?'diff-up':tD<0?'diff-down':'diff-same';
  const tArrow=tD>0?'▲':tD<0?'▼':'→';
  const tPct=tA>0?` (${tD>0?'+':''}${Math.round(tD/tA*100)}%)`:'';

  document.getElementById('weekCompareGrid').innerHTML = `
    <thead><tr>
      <th>단말기</th>
      <th>${wA.week_label}</th>
      <th>${wB.week_label}</th>
      <th>변화량</th>
    </tr></thead>
    <tbody>
      ${rows}
      <tr>
        <td>합계</td>
        <td><span class="cmp-count">${tA}건</span></td>
        <td><span class="cmp-count">${tB}건</span></td>
        <td><span class="cmp-diff ${tCls}">${tArrow} ${Math.abs(tD)}건${tPct}</span></td>
      </tr>
    </tbody>`;

  // 단말기 셀렉트 + 컨테이너 초기 렌더
  const activeTabs = TAB_ORDER.filter(t => (wA.tab_counts[t]||0)+(wB.tab_counts[t]||0) > 0);
  const tabOpts = `<option value="__all__">전체 합산</option>` +
    activeTabs.map(t => `<option value="${t}">${t}</option>`).join('');
  document.getElementById('faultChangeWrap').innerHTML = `
    <div style="display:flex;align-items:center;gap:10px;margin:16px 0 12px;flex-wrap:wrap">
      <span style="font-size:13px;font-weight:700;color:var(--gray-dark)">장애유형 변화</span>
      <select class="month-sel" id="faultTabSel" onchange="renderFaultChange()">${tabOpts}</select>
    </div>
    <div id="faultChangeContent"></div>`;

  renderFaultChange();
}

function renderFaultChange() {
  const idxA = parseInt(document.getElementById('weekSelA').value);
  const idxB = parseInt(document.getElementById('weekSelB').value);
  const wA = confirmedWeeksData[idxA], wB = confirmedWeeksData[idxB];
  const sel = document.getElementById('faultTabSel');
  const selectedTab = sel ? sel.value : '__all__';

  function getFaults(week) {
    if (selectedTab === '__all__') {
      const map = {};
      for (const faults of Object.values(week.tab_faults||{}))
        for (const [ft, cnt] of Object.entries(faults))
          map[ft] = (map[ft]||0) + cnt;
      return map;
    }
    return week.tab_faults?.[selectedTab] || {};
  }

  const fA=getFaults(wA), fB=getFaults(wB);
  const allFaultTypes = new Set([...Object.keys(fA), ...Object.keys(fB)]);

  const increased=[], decreased=[], newFaults=[], goneFaults=[];
  allFaultTypes.forEach(ft => {
    const a=fA[ft]||0, b=fB[ft]||0;
    if (a===0 && b>0)      newFaults.push({ft,a,b,d:b});
    else if (b===0 && a>0) goneFaults.push({ft,a,b,d:-a});
    else if (b>a)          increased.push({ft,a,b,d:b-a});
    else if (b<a)          decreased.push({ft,a,b,d:b-a});
  });
  increased.sort((x,y)=>y.d-x.d);
  decreased.sort((x,y)=>x.d-y.d);
  newFaults.sort((x,y)=>y.b-x.b);
  goneFaults.sort((x,y)=>y.a-x.a);

  function faultItemHtml(items, type) {
    if (!items.length) return `<div style="color:var(--gray);font-size:12px;padding:8px 0">없음</div>`;
    return items.slice(0,10).map(({ft,a,b,d}) => {
      let deltaHtml;
      if (type==='new')       deltaHtml=`<span class="fault-delta new">NEW +${b}건</span>`;
      else if (type==='gone') deltaHtml=`<span class="fault-delta gone">GONE -${a}건</span>`;
      else if (d>0)           deltaHtml=`<span class="fault-delta up">▲ +${d}건 (${a}→${b})</span>`;
      else                    deltaHtml=`<span class="fault-delta down">▼ ${d}건 (${a}→${b})</span>`;
      return `<div class="fault-item"><span class="fault-name">${ft}</span>${deltaHtml}</div>`;
    }).join('');
  }

  const noChange = !increased.length && !decreased.length && !newFaults.length && !goneFaults.length;
  document.getElementById('faultChangeContent').innerHTML = noChange
    ? `<div style="color:var(--gray);font-size:13px;padding:12px 0">변화 없음</div>`
    : `<div class="fault-change-grid">
        <div class="fault-change-box up">
          <div class="fault-change-title up">▲ 증가${newFaults.length?` &nbsp;&middot;&nbsp; 🆕 신규 ${newFaults.length}종`:''}</div>
          ${faultItemHtml(increased,'up')}
          ${newFaults.length?`<div style="border-top:1px dashed #F5C0CF;margin:8px 0"></div>${faultItemHtml(newFaults,'new')}`:''}
        </div>
        <div class="fault-change-box down">
          <div class="fault-change-title down">▼ 감소${goneFaults.length?` &nbsp;&middot;&nbsp; 소멸 ${goneFaults.length}종`:''}</div>
          ${faultItemHtml(decreased,'down')}
          ${goneFaults.length?`<div style="border-top:1px dashed #A8D8B4;margin:8px 0"></div>${faultItemHtml(goneFaults,'gone')}`:''}
        </div>
      </div>`;
}

// ── 단말기별 장애유형 ────────────────────────────────
let ftAllData = null;
// 탭 → 그룹 매핑 (뷰 이름 결정용)
const FT_TAB_META = {
  b_series: { cls: 'b-tab',  views: ['접수오류유형','현장처리유형'] },
  airport:  { cls: 'ap-tab', views: ['접수오류유형','현장처리유형'] },
  regional: { cls: 'rg-tab', views: ['단말기접수유형','단말기현장처리'] },
};
const ftState = { groupKey: null, tab: null, view: null };

// 전체 탭 플랫 리스트 반환: [{tab, groupKey, cls}]
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

function loadFaultTypeSection() {
  fetch('/api/fault-type-stats', { cache: 'no-store' })
    .then(r => r.json())
    .then(data => {
      if (!data.ok) return;
      ftAllData = data;
      const flat = ftFlatTabs();
      if (!flat.length) return;
      ftState.groupKey = flat[0].groupKey;
      ftState.tab = flat[0].tab;
      ftState.view = FT_TAB_META[flat[0].groupKey].views[0];
      document.getElementById('faultTypeSection').style.display = '';
      renderFtTabGroup();
      renderFtViewToggle();
      renderFtTable();
    })
    .catch(e => console.error('장애유형 데이터 오류:', e));
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

function renderFtTable() {
  const tbl = document.getElementById('ftTable');
  const gd = ftAllData?.[ftState.groupKey];
  if (!gd || !gd.tabs[ftState.tab]) { tbl.innerHTML = ''; return; }

  const td = gd.tabs[ftState.tab];
  const months = gd.months;
  const typeData = td[ftState.view];
  if (!typeData || !typeData.types?.length) {
    tbl.innerHTML = `<tr><td class="empty-msg">데이터 없음</td></tr>`; return;
  }

  const { types, by_month, total_by_type } = typeData;
  const grandTotal = months.reduce((s,m) => s+(td.by_month[m]||0), 0);
  const monthTotals = {};
  months.forEach(m => {
    monthTotals[m] = types.reduce((s,t) => s+(by_month[t]?.[m]||0), 0);
  });

  const typeTotal = types.reduce((s,t) => s+(total_by_type[t]||0), 0);

  const rows = types.map(type => {
    const cells = months.map(m => {
      const v = by_month[type]?.[m] || 0;
      return `<td>${v > 0 ? v.toLocaleString()+'건' : '<span style="color:var(--border)">-</span>'}</td>`;
    }).join('');
    const tot = total_by_type[type] || 0;
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
