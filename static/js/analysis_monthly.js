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
function getTabClass(t){ return t.startsWith("서울")?"seoul":t.startsWith("공항")?"airport":"regional"; }

let allData = null;
let activeCharts = {};
let activeTabs = new Set(TAB_ORDER);
let statsYear = 'all'; // 통계 테이블 전용 연도 필터

// ─── 연도 추출 헬퍼 (key 형식: "202506" 또는 "2025-06") ───
function keyToYear(k) {
  return k.includes('-') ? k.split('-')[0] : k.substring(0, 4);
}

function allKeys() {
  if (!allData) return [];
  return Object.keys(allData.all_stats).sort();
}

// ─── 초기 로드 ──────────────────────────────────────
function loadDashboard(isReload) {
  fetch('/api/analysis-data', { cache: 'no-store' })
    .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); })
    .then(data => {
      if (!data.ok) { showError('데이터 오류: ' + (data.error || '알 수 없는 오류')); return; }
      allData = data;
      document.getElementById('loadingBar').style.display = 'none';
      document.getElementById('dashContent').style.display = 'block';

      try { initTabChecks(); } catch(e) { console.error('탭 초기화 오류:', e); }
      try { renderStatsYearSeg(); renderStatsTable(); } catch(e) { console.error('테이블 오류:', e); }
      try { renderDownloads(data.all_stats || {}); } catch(e) { console.error('다운로드 오류:', e); }
      try { initCompareSelects(); } catch(e) { console.error('셀렉트 오류:', e); }
      try { renderCompare(); } catch(e) { console.error('비교 오류:', e); }

      if (isReload) {
        try {
          if (activeCharts.count) { activeCharts.count.destroy(); activeCharts.count = null; }
          if (activeCharts.rate)  { activeCharts.rate.destroy();  activeCharts.rate  = null; }
          renderCharts();
        } catch(e) { console.error('차트 재생성 오류:', e); }
      } else {
        try { renderCharts(); } catch(e) {
          console.error('차트 오류:', e);
          document.getElementById('countChart').parentElement.innerHTML = '<p style="color:var(--gray);font-size:12px;padding:20px">차트를 불러올 수 없습니다.</p>';
          document.getElementById('rateChart').parentElement.innerHTML = '<p style="color:var(--gray);font-size:12px;padding:20px">차트를 불러올 수 없습니다.</p>';
        }
      }
    })
    .catch(e => { showError('서버 연결 오류: ' + e.message); });
}

window.addEventListener('DOMContentLoaded', () => loadDashboard(false));
document.addEventListener('visibilitychange', () => { if (!document.hidden && allData) loadDashboard(true); });

function showError(msg) {
  const bar = document.getElementById('loadingBar');
  bar.textContent = msg;
  bar.style.color = 'var(--primary)';
}

// ─── 탭 필터 (차트용) ────────────────────────────────
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

// ─── 차트 (전체 데이터 사용) ──────────────────────────
function buildDatasets(field) {
  const keys = allKeys();
  return TAB_ORDER.map(tab => {
    const ds = allData.chart_datasets[tab] || {};
    const vals = keys.map((k, i) => {
      const v = (ds[field] || [])[i];
      return v == null ? 0 : Number(v);
    });
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
  const labels = allKeys().map(k => allData.all_stats[k].label);

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
      legend: { position:'bottom', labels:{ font:{size:11}, boxWidth:12, padding:8 }, onClick:legendClickHandler },
      tooltip: { mode:'index', intersect:false },
    },
    interaction: { mode:'nearest', axis:'x', intersect:false },
  };

  activeCharts.count = new Chart(document.getElementById('countChart'), {
    type:'line',
    data: { labels, datasets: buildDatasets('counts') },
    options: { ...commonOpts, scales:{ x:{ticks:{font:{size:10}}}, y:{beginAtZero:true,ticks:{font:{size:10}}} } },
  });
  activeCharts.rate = new Chart(document.getElementById('rateChart'), {
    type:'line',
    data: { labels, datasets: buildDatasets('rates') },
    options: { ...commonOpts, scales:{ x:{ticks:{font:{size:10}}}, y:{beginAtZero:true,ticks:{callback:v=>v+'%',font:{size:10}}} } },
  });
}

function updateCharts() {
  if (!allData || !activeCharts.count) return;
  [activeCharts.count, activeCharts.rate].forEach(chart => {
    chart.data.datasets.forEach(ds => { ds.hidden = !activeTabs.has(ds.label); });
    chart.update();
  });
}

// ─── 통계 테이블 연도 필터 ────────────────────────────
function renderStatsYearSeg() {
  const years = [...new Set(allKeys().map(keyToYear))].sort();
  const btns = [{ val: 'all', label: '전체' }, ...years.map(y => ({ val: y, label: y + '년' }))];
  document.getElementById('statsYearSeg').innerHTML =
    `<div class="year-seg">${btns.map(b =>
      `<button class="year-seg-btn ${b.val===statsYear?'active':''}" onclick="setStatsYear('${b.val}')">${b.label}</button>`
    ).join('')}</div>`;
}

function setStatsYear(year) {
  statsYear = year;
  document.querySelectorAll('#statsYearSeg .year-seg-btn').forEach(b => {
    b.classList.toggle('active', b.textContent === (year === 'all' ? '전체' : year + '년'));
  });
  renderStatsTable();
}

// ─── 전체 통계 테이블 (연도별 세로 분리) ──────────────
function renderStatsTable() {
  const all_stats = allData.all_stats;
  const keys = allKeys();
  const years = [...new Set(keys.map(keyToYear))].sort();
  const targetYears = statsYear === 'all' ? years : [statsYear];
  const wrap = document.getElementById('statsTableWrap');

  wrap.innerHTML = targetYears.map((y, yi) => {
    const yKeys = keys.filter(k => keyToYear(k) === y);
    const monthLabels = yKeys.map(k => {
      // "2025년 6월" → "6월" 로 줄여서 표시 (연도는 헤더에 이미 표시)
      return all_stats[k].label.replace(/^\d{4}년 /, '');
    });

    const headHtml =
      `<tr><th>단말기</th>` + monthLabels.map(l => `<th colspan="2">${l}</th>`).join('') + `</tr>` +
      `<tr><th></th>` + monthLabels.map(() => `<th>건수</th><th>장애율</th>`).join('') + `</tr>`;

    const bodyHtml = TAB_ORDER.map(tab => {
      const cells = yKeys.map(k => {
        const d = ((all_stats[k].tabs || {})[tab]) || {};
        const cnt = d.count != null ? d.count : '-';
        const rate = d.fault_rate != null ? d.fault_rate + '%' : '-';
        const rateCls = d.fault_rate != null ? (d.fault_rate >= 1 ? 'rate-bad' : 'rate-ok') : '';
        return `<td>${cnt}</td><td class="rate-cell ${rateCls}">${rate}</td>`;
      }).join('');
      return `<tr><td style="text-align:center"><span class="tab-label-analysis ${getTabClass(tab)}">${tab}</span></td>${cells}</tr>`;
    }).join('');

    const marginTop = yi > 0 ? 'margin-top:24px' : '';
    return `
      <div style="${marginTop}">
        <div style="font-size:13px;font-weight:800;color:var(--gray-dark);margin-bottom:8px;padding-left:2px">${y}년</div>
        <div style="overflow-x:auto">
          <table class="stats-month-table">
            <thead>${headHtml}</thead>
            <tbody>${bodyHtml}</tbody>
          </table>
        </div>
      </div>`;
  }).join('');
}

// ─── 월 비교 셀렉트 (전체 데이터 기준) ──────────────
function initCompareSelects() {
  const keys = allKeys();
  const all_stats = allData.all_stats;
  const opts = keys.map((k, i) => `<option value="${i}">${all_stats[k].label}</option>`).join('');
  document.getElementById('monthSelA').innerHTML = opts;
  document.getElementById('monthSelB').innerHTML = opts;
  if (keys.length >= 2) {
    document.getElementById('monthSelA').value = String(keys.length - 2);
    document.getElementById('monthSelB').value = String(keys.length - 1);
  }
}

function renderCompare() {
  if (!allData) return;
  const keys = allKeys();
  const all_stats = allData.all_stats;
  const idxA = parseInt(document.getElementById('monthSelA').value || '0');
  const idxB = parseInt(document.getElementById('monthSelB').value || '1');
  const statA = all_stats[keys[idxA]];
  const statB = all_stats[keys[idxB]];
  if (!statA || !statB) return;

  const maxCount = Math.max(...TAB_ORDER.map(t => Math.max((statA.tabs[t]||{}).count||0, (statB.tabs[t]||{}).count||0)), 1);

  function cellA(tab) {
    const d = (statA.tabs||{})[tab]||{};
    const cnt = d.count||0, rate = d.fault_rate;
    const rateCls = rate==null?'':rate>=1?'bad':'ok';
    return `<td><span class="cmp-count">${cnt}건</span><span class="cmp-rate ${rateCls}">장애율 ${rate!=null?rate+'%':'-'}</span></td>`;
  }
  function cellB(tab) {
    const d = (statB.tabs||{})[tab]||{};
    const cnt = d.count||0, rate = d.fault_rate;
    const rateCls = rate==null?'':rate>=1?'bad':'ok';
    return `<td><span class="cmp-count">${cnt}건</span><span class="cmp-rate ${rateCls}">장애율 ${rate!=null?rate+'%':'-'}</span></td>`;
  }
  function cellDiff(tab) {
    const cA=(statA.tabs[tab]||{}).count||0, cB=(statB.tabs[tab]||{}).count||0;
    const diff=cB-cA, cls=diff>0?'diff-up':diff<0?'diff-down':'diff-same';
    const arrow=diff>0?'▲':diff<0?'▼':'→';
    const barW=Math.round(Math.abs(diff)/maxCount*60);
    const barColor=diff>0?'var(--primary)':diff<0?'var(--green)':'var(--gray)';
    const barHtml=diff!==0?`<span class="cmp-bar" style="width:${barW}px;background:${barColor}"></span>`:'';
    const pct=cA>0?` (${diff>0?'+':''}${Math.round(diff/cA*100)}%)`:'';
    return `<td><span class="cmp-diff ${cls}">${arrow} ${Math.abs(diff)}건${pct}</span>${barHtml}</td>`;
  }

  const tA=Object.values(statA.tabs).reduce((s,d)=>s+(d.count||0),0);
  const tB=Object.values(statB.tabs).reduce((s,d)=>s+(d.count||0),0);
  const tDiff=tB-tA, tCls=tDiff>0?'diff-up':tDiff<0?'diff-down':'diff-same';
  const tArrow=tDiff>0?'▲':tDiff<0?'▼':'→';
  const tPct=tA>0?` (${tDiff>0?'+':''}${Math.round(tDiff/tA*100)}%)`:'';

  const rows = TAB_ORDER.map(tab => {
    const cA=(statA.tabs[tab]||{}).count||0, cB=(statB.tabs[tab]||{}).count||0;
    if (cA===0 && cB===0) return '';
    return `<tr><td><span class="tab-label-analysis ${getTabClass(tab)}">${tab}</span></td>${cellA(tab)}${cellB(tab)}${cellDiff(tab)}</tr>`;
  }).join('');

  document.getElementById('compareGrid').innerHTML = `
    <thead><tr><th>단말기</th><th>${statA.label}</th><th>${statB.label}</th><th>변화량</th></tr></thead>
    <tbody>${rows}
      <tr><td>합계</td><td><span class="cmp-count">${tA}건</span></td><td><span class="cmp-count">${tB}건</span></td>
        <td><span class="cmp-diff ${tCls}">${tArrow} ${Math.abs(tDiff)}건${tPct}</span></td></tr>
    </tbody>`;
}

// ─── 다운로드 ────────────────────────────────────────
function renderDownloads(all_stats) {
  const row = document.getElementById('downloadRow');
  const confirmedKeys = Object.keys(all_stats).sort().filter(k => all_stats[k].confirmed);
  row.innerHTML = confirmedKeys.map(k => {
    const fn = all_stats[k].filename;
    const label = all_stats[k].label;
    return `<button class="btn-dl confirmed" onclick="downloadFile('${encodeURIComponent(fn)}')">&#128274; ${label} &#11015;</button>`;
  }).join('') || '<p style="color:var(--gray);font-size:12px">다운로드 가능한 파일이 없습니다.</p>';
}

function downloadFile(encoded) {
  window.location.href = '/api/download-monthly/' + encoded;
}
