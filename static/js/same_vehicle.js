let svAllRows = [];
let svCurrentTab = '전체';

function doSearch() {
  const btn = document.getElementById('btnSearch');
  const loading = document.getElementById('svLoading');
  const body = document.getElementById('svTableBody');

  btn.disabled = true;
  loading.style.display = 'block';
  body.innerHTML = '';
  document.getElementById('svEmptyMsg').style.display = 'none';
  document.getElementById('resultSection').style.display = 'block';

  const payload = {
    date_from: document.getElementById('dateFrom').value,
    date_to: document.getElementById('dateTo').value,
    terminal: document.getElementById('terminalSel').value,
    min_count: parseInt(document.getElementById('minCountSel').value),
  };

  fetch('/api/same-vehicle-stats', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
    .then(r => r.json())
    .then(data => {
      if (!data.ok) throw new Error(data.error || '오류');
      svAllRows = data.rows || [];
      document.getElementById('cardTotal').textContent = data.total.toLocaleString();
      document.getElementById('cardSame').textContent = data.same_fault.toLocaleString();
      document.getElementById('cardDiff').textContent = data.diff_fault.toLocaleString();
      renderTable();
    })
    .catch(e => {
      showToast('오류: ' + e.message, 'error');
    })
    .finally(() => {
      btn.disabled = false;
      loading.style.display = 'none';
    });
}

function setTab(el) {
  document.querySelectorAll('.sv-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  svCurrentTab = el.dataset.type;
  renderTable();
}

function renderTable() {
  const body = document.getElementById('svTableBody');
  const emptyMsg = document.getElementById('svEmptyMsg');

  const rows = svCurrentTab === '전체'
    ? svAllRows
    : svAllRows.filter(r => r.유형 === svCurrentTab);

  if (!rows.length) {
    body.innerHTML = '';
    emptyMsg.style.display = 'block';
    return;
  }
  emptyMsg.style.display = 'none';

  // 같은 차량 첫 행 표시용 추적
  let prevCar = null;
  let html = '';
  for (const r of rows) {
    const isFirst = r.차량번호 !== prevCar;
    prevCar = r.차량번호;
    const rowClass = r.유형 === '동일장애'
      ? (isFirst ? 'row-same row-first-same' : 'row-same')
      : (isFirst ? 'row-diff row-first-diff' : 'row-diff');

    const badge = r.유형 === '동일장애'
      ? `<span class="sv-badge badge-same">동일장애</span>`
      : `<span class="sv-badge badge-diff">다중장애</span>`;

    const cntCell = isFirst
      ? `<span class="cnt-badge">${r.차량건수}</span>`
      : '';

    html += `<tr class="${rowClass}">
      <td>${esc(r.차량번호)}</td>
      <td style="text-align:center">${cntCell}</td>
      <td style="white-space:nowrap">${esc(r.날짜)}</td>
      <td>${esc(r.배정부서)}</td>
      <td style="text-align:center"><span class="week-badge" style="font-size:11px">${esc(r.단말기구분)}</span></td>
      <td>${esc(r.접수오류유형)}</td>
      <td>${esc(r.교통사업자명)}</td>
      <td>${esc(r.처리유형)}</td>
      <td>${esc(r.처리자)}</td>
      <td style="white-space:nowrap">${esc(r.처리완료일시)}</td>
      <td style="text-align:center">${badge}</td>
    </tr>`;
  }
  body.innerHTML = html;
}

function esc(v) {
  if (!v) return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

let toastTimer;
function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast' + (type ? ' ' + type : '') + ' show';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
}

// 오늘 날짜로 dateTo 초기화
(function() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  document.getElementById('dateTo').value = `${yyyy}-${mm}-${dd}`;
})();
