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
window.addEventListener('DOMContentLoaded', () => {
  fetch('/api/analysis-data')
    .then(r => {
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      return r.json();
    })
    .then(data => {
      if (!data.ok) {
        showError('데이터 오류: ' + (data.error || '알 수 없는 오류'));
        return;
      }
      allData = data;

      // 먼저 UI 표시
      document.getElementById('loadingBar').style.display = 'none';
      document.getElementById('dashContent').style.display = 'block';

      // 차트 없는 콘텐츠 먼저 렌더 (에러 없이 보여줄 것들)
      const sortedKeys = Object.keys(data.all_stats || {}).sort();

      try { initTabChecks(); } catch(e) { console.error('탭 초기화 오류:', e); }
      try { renderComments(data.comments || []); } catch(e) { console.error('코멘트 오류:', e); }
      try { initCommentSelects(sortedKeys, data.all_stats || {}); } catch(e) { console.error('코멘트 셀렉트 오류:', e); }
      try { renderStatsTable(data.all_stats || {}); } catch(e) { console.error('테이블 오류:', e); }
      try { renderDownloads(data.all_stats || {}); } catch(e) { console.error('다운로드 오류:', e); }
      try { initCompareSelects(data.labels || []); } catch(e) { console.error('셀렉트 오류:', e); }
      try { renderCompare(); } catch(e) { console.error('비교 오류:', e); }

      // 차트는 마지막에 (실패해도 나머지는 보임)
      try { renderCharts(); } catch(e) {
        console.error('차트 오류:', e);
        document.getElementById('countChart').parentElement.innerHTML =
          '<p style="color:var(--gray);font-size:12px;padding:20px">차트를 불러올 수 없습니다.</p>';
        document.getElementById('rateChart').parentElement.innerHTML =
          '<p style="color:var(--gray);font-size:12px;padding:20px">차트를 불러올 수 없습니다.</p>';
      }
    })
    .catch(e => {
      showError('서버 연결 오류: ' + e.message);
    });
});

function showError(msg) {
  document.getElementById('loadingBar').textContent = msg;
  document.getElementById('loadingBar').style.color = 'var(--primary)';
}

// ─── 탭 필터 ─────────────────────────────────────────
function initTabChecks() {
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
  return TAB_ORDER
    .filter(tab => activeTabs.has(tab))
    .map(tab => {
      const ds = allData.chart_datasets[tab] || {};
      const vals = (ds[field] || []).map(v => v == null ? null : Number(v));
      return {
        label: tab, data: vals,
        borderColor: TAB_COLORS[tab] || '#888',
        backgroundColor: (TAB_COLORS[tab] || '#888') + '33',
        borderWidth: 2, pointRadius: 3, tension: 0.3, fill: false,
      };
    });
}

function renderCharts() {
  if (typeof Chart === 'undefined') throw new Error('Chart.js not loaded');

  const commonOpts = {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { position: 'bottom', labels: { font: { size: 11 }, boxWidth: 12, padding: 8 } },
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
      datasets: buildDatasets('rates').map(ds => ({ ...ds, spanGaps: false })),
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
  activeCharts.count.data.datasets = buildDatasets('counts');
  activeCharts.count.update();
  activeCharts.rate.data.datasets = buildDatasets('rates').map(ds => ({ ...ds, spanGaps: false }));
  activeCharts.rate.update();
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

// ─── 코멘트 기간 셀렉트 ──────────────────────────────
function initCommentSelects(sortedKeys, all_stats) {
  const opts = sortedKeys.map(k => `<option value="${k}">${all_stats[k].label}</option>`).join('');
  document.getElementById('commentSelA').innerHTML = opts;
  document.getElementById('commentSelB').innerHTML = opts;
  if (sortedKeys.length >= 2) {
    document.getElementById('commentSelA').value = sortedKeys[sortedKeys.length - 2];
    document.getElementById('commentSelB').value = sortedKeys[sortedKeys.length - 1];
  }
}

function refreshComments() {
  if (!allData) return;
  const fromKey = document.getElementById('commentSelA').value;
  const toKey = document.getElementById('commentSelB').value;
  const list = document.getElementById('commentList');
  list.innerHTML = '<li class="comment-item"><span class="comment-text" style="color:var(--gray)">불러오는 중...</span></li>';
  fetch(`/api/analysis-comments?from=${fromKey}&to=${toKey}`)
    .then(r => r.json())
    .then(data => {
      if (data.ok) renderComments(data.comments || []);
      else list.innerHTML = '<li class="comment-item"><span class="comment-text" style="color:var(--gray)">코멘트 로드 실패</span></li>';
    })
    .catch(() => {
      list.innerHTML = '<li class="comment-item"><span class="comment-text" style="color:var(--gray)">서버 오류</span></li>';
    });
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

// ─── 코멘트 ──────────────────────────────────────────
const CSTYLE = {
  summary_bad:  {cls:'bad',    lbl:'전체▲'},
  summary_good: {cls:'good',   lbl:'전체▼'},
  summary_same: {cls:'info',   lbl:'전체→'},
  tab_bad:      {cls:'bad',    lbl:'증가▲'},
  tab_good:     {cls:'good',   lbl:'감소▼'},
  tab_same:     {cls:'info',   lbl:'유지→'},
  alert:        {cls:'alert',  lbl:'경보⚠'},
  rate_info:    {cls:'info',   lbl:'장애율'},
  new_bad:      {cls:'new',    lbl:'신규▲'},
  new_good:     {cls:'new',    lbl:'신규▼'},
  new_same:     {cls:'new',    lbl:'신규→'},
  divider:      {cls:'divider',lbl:''},
};

function renderComments(comments) {
  const list = document.getElementById('commentList');
  if (!comments.length) {
    list.innerHTML = '<li class="comment-item"><span class="comment-text" style="color:var(--gray)">분석 데이터가 충분하지 않습니다.</span></li>';
    return;
  }
  list.innerHTML = comments.map(c => {
    const s = CSTYLE[c.type] || {cls:'info', lbl: c.tag || ''};
    if (c.type === 'divider') {
      return `<li class="comment-item divider-row"><span class="comment-text">${c.text}</span></li>`;
    }
    return `<li class="comment-item">
      <span class="c-tag ${s.cls}">${s.lbl || c.tag}</span>
      <span class="comment-text">${c.text}</span>
    </li>`;
  }).join('');
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
