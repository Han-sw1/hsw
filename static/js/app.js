let selectedFiles = [];
let resultFilename = null;
let previewData = {};
let statsData = {};
let detectedWeeks = [];

const TAB_ORDER = [
  "서울 B800", "서울 B700", "서울 B710", "공항 B620",
  "대전 B650", "세종 B500", "제주 B400", "포항 B800",
  "상주,영주,예천 B400", "안동 B520D", "김해 B600"
];

const MODE_DESC = {
  daily:   "새로운 장애 데이터를 추가합니다. 이미 있는 접수번호는 중복 추가하지 않습니다.",
  weekly:  "선택한 주차의 기존 데이터를 삭제하고 새 데이터로 교체합니다. (목요일 주차 마감용)",
  monthly: "해당 시트의 전체 데이터를 새 데이터로 교체합니다. (월 마감·확정용)",
};

function getTabClass(tab) {
  if (tab.startsWith("서울")) return "seoul";
  if (tab.startsWith("공항")) return "airport";
  return "regional";
}

// ─── 설정 모달 ───────────────────────────────────────
document.getElementById('btnSettings').addEventListener('click', () => {
  document.getElementById('settingsModal').classList.add('open');
  // 서버에 저장된 기준파일 현황 로드
  fetch('/api/reference-files').then(r => r.json()).then(data => {
    const cs = document.getElementById('criteriaStatus');
    const ks = document.getElementById('citsStatus');
    if (data.criteria) {
      cs.textContent = '✔ ' + data.criteria;
      cs.className = 'ref-status ok';
    } else {
      cs.textContent = '미등록 — 파일을 업로드하세요';
      cs.className = 'ref-status missing';
    }
    if (data.cits) {
      ks.textContent = '✔ ' + data.cits;
      ks.className = 'ref-status ok';
    } else {
      ks.textContent = '미등록 — 파일을 업로드하세요';
      ks.className = 'ref-status missing';
    }
  }).catch(() => {});
  // 로컬 경로 fallback
  fetch('/api/config').then(r => r.json()).then(cfg => {
    document.getElementById('criteriaPath').value = cfg.criteria_path || '';
    document.getElementById('citsPath').value = cfg.cits_path || '';
  }).catch(() => {});
});

function uploadReferenceFile(inputEl, type, statusEl) {
  const file = inputEl.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  fd.append('type', type);
  statusEl.textContent = '업로드 중...';
  statusEl.className = 'ref-status';
  fetch('/api/upload-reference', { method: 'POST', body: fd })
    .then(r => r.json())
    .then(data => {
      if (data.ok) {
        statusEl.textContent = '✔ ' + file.name;
        statusEl.className = 'ref-status ok';
        showToast(file.name + ' 업로드 완료', 'success');
      } else {
        statusEl.textContent = '업로드 실패: ' + (data.error || '오류');
        statusEl.className = 'ref-status missing';
      }
    })
    .catch(() => {
      statusEl.textContent = '업로드 실패';
      statusEl.className = 'ref-status missing';
    });
  inputEl.value = '';
}

document.getElementById('criteriaFileInput').addEventListener('change', function() {
  uploadReferenceFile(this, 'criteria', document.getElementById('criteriaStatus'));
});
document.getElementById('citsFileInput').addEventListener('change', function() {
  uploadReferenceFile(this, 'cits', document.getElementById('citsStatus'));
});
['closeSettings', 'cancelSettings'].forEach(id =>
  document.getElementById(id).addEventListener('click', closeSettings)
);
document.getElementById('settingsModal').addEventListener('click', e => {
  if (e.target === e.currentTarget) closeSettings();
});
function closeSettings() {
  document.getElementById('settingsModal').classList.remove('open');
}
document.getElementById('saveSettings').addEventListener('click', () => {
  const cfg = {
    criteria_path: document.getElementById('criteriaPath').value.trim(),
    cits_path: document.getElementById('citsPath').value.trim(),
  };
  fetch('/api/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(cfg),
  }).then(() => {
    closeSettings();
    showToast('설정이 저장되었습니다.', 'success');
  });
});

// ─── 파일 선택 ───────────────────────────────────────
const fileInput = document.getElementById('fileInput');
const dropZone  = document.getElementById('dropZone');

fileInput.addEventListener('change', e => {
  addFiles(Array.from(e.target.files));
  fileInput.value = '';
});
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const valid = Array.from(e.dataTransfer.files).filter(f =>
    f.name.endsWith('.xls') || f.name.endsWith('.xlsx')
  );
  if (!valid.length) return showToast('.xls 또는 .xlsx 파일만 가능합니다.', 'error');
  addFiles(valid);
});
dropZone.addEventListener('click', e => {
  if (e.target.tagName !== 'BUTTON') fileInput.click();
});

function addFiles(files) {
  const existing = new Set(selectedFiles.map(f => f.name));
  files.forEach(f => { if (!existing.has(f.name)) selectedFiles.push(f); });
  renderFileList();
  document.getElementById('btnProcess').disabled = selectedFiles.length === 0;
}

function removeFile(name) {
  selectedFiles = selectedFiles.filter(f => f.name !== name);
  renderFileList();
  document.getElementById('btnProcess').disabled = selectedFiles.length === 0;
  if (!selectedFiles.length) {
    ['resultCard','previewCard','statsCard','uploadCompareCard','monthlyCard'].forEach(id => {
      document.getElementById(id).style.display = 'none';
    });
  }
}

function renderFileList(confirmed = false) {
  const list = document.getElementById('fileList');
  list.innerHTML = selectedFiles.map(f => `
    <div class="file-item">
      <span class="file-icon">&#128196;</span>
      <span class="file-name">${f.name}</span>
      <span class="file-type-badge ${confirmed ? 'badge-confirmed' : 'badge-pending'}"
            data-name="${f.name}">${confirmed ? '&#10003; 확인됨' : '감지 전'}</span>
      <button class="btn-remove-file" onclick="removeFile('${f.name.replace(/'/g, "\\'")}')">&#10005;</button>
    </div>
  `).join('');
}

// ─── 처리 ──────────────────────────────────────────
document.getElementById('btnProcess').addEventListener('click', async () => {
  if (!selectedFiles.length) return;
  document.getElementById('guideCard').style.display = 'none';
  showLoading(true, '처리 중입니다...');

  const formData = new FormData();
  selectedFiles.forEach(f => formData.append('files', f));

  try {
    const res = await fetch('/api/process', { method: 'POST', body: formData });
    const data = await res.json();

    if (!data.ok) {
      showToast(data.error || '처리 오류', 'error');
      return;
    }

    resultFilename = data.filename;
    previewData = data.previews || {};
    statsData = data.stats || {};
    detectedWeeks = data.weeks || [];

    renderFileList(true);
    renderSummary(data.summary || {});
    renderTabButtons();
    renderStats(statsData);
    renderUploadComparison(data.upload_comments || [], data.week_comments || []);
    setupMonthlySection();

    document.getElementById('resultCard').style.display = 'block';
    document.getElementById('previewCard').style.display = 'block';
    document.getElementById('statsCard').style.display = 'block';
    document.getElementById('monthlyCard').style.display = 'block';
    document.getElementById('insertReport').style.display = 'none';

    const firstTab = TAB_ORDER.find(t => previewData[t] && previewData[t].count > 0) || TAB_ORDER[0];
    selectTab(firstTab);

    const total = Object.values(data.summary || {}).reduce((a, b) => a + b, 0);
    showToast(`처리 완료! 총 ${total}건`, 'success');
    document.getElementById('resultCard').scrollIntoView({ behavior: 'smooth' });

  } catch (e) {
    showToast('서버 오류: ' + e.message, 'error');
  } finally {
    showLoading(false);
  }
});

// ─── 기존 처리결과 다운로드 ───────────────────────────
document.getElementById('btnDownload').addEventListener('click', () => {
  if (!resultFilename) return;
  window.location.href = '/api/download/' + encodeURIComponent(resultFilename);
});

// ─── 요약 렌더 ────────────────────────────────────
function renderSummary(summary) {
  const grid = document.getElementById('summaryGrid');
  grid.innerHTML = TAB_ORDER.map(tab => {
    const count = summary[tab] ?? 0;
    const isEmpty = count === 0;
    const cls = getTabClass(tab);
    return `
      <div class="summary-item ${isEmpty ? 'empty-tab' : cls}" onclick="selectTab('${tab.replace(/'/g,"\\'")}')">
        <div class="summary-tab">${tab}</div>
        <div class="summary-count">${isEmpty ? '없음' : count}</div>
        ${isEmpty ? '' : '<div class="summary-label">건</div>'}
      </div>`;
  }).join('');
}

// ─── 탭 버튼 ─────────────────────────────────────
function renderTabButtons() {
  const btns = document.getElementById('tabButtons');
  btns.innerHTML = TAB_ORDER.map(tab => {
    const cls = getTabClass(tab);
    const count = previewData[tab] ? previewData[tab].count : 0;
    return `<button class="tab-btn ${cls}" data-tab="${tab}" onclick="selectTab('${tab.replace(/'/g,"\\'")}')">
      ${tab} ${count > 0 ? `<span style="opacity:.7">(${count})</span>` : ''}
    </button>`;
  }).join('');
}

let currentTab = null;
function selectTab(tab) {
  currentTab = tab;
  document.querySelectorAll('.summary-item').forEach((el, i) => {
    el.classList.toggle('active', TAB_ORDER[i] === tab);
  });
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tab);
  });

  const info = previewData[tab];
  const table = document.getElementById('previewTable');
  const emptyNotice = document.getElementById('emptyNotice');
  const head = document.getElementById('previewHead');
  const body = document.getElementById('previewBody');

  if (!info || !info.cols || info.cols.length === 0) {
    table.style.display = 'none';
    emptyNotice.style.display = 'block';
    return;
  }

  table.style.display = 'table';
  emptyNotice.style.display = 'none';

  const HIGHLIGHT = new Set(['날짜','cits','월','주차','장애접수일']);
  head.innerHTML = '<tr>' + info.cols.map(c => `<th>${c}</th>`).join('') + '</tr>';
  body.innerHTML = info.rows.map(row =>
    '<tr>' + info.cols.map(c =>
      `<td class="${HIGHLIGHT.has(c) ? 'col-highlight' : ''}">${row[c] ?? ''}</td>`
    ).join('') + '</tr>'
  ).join('');
}

// ─── 집계 현황 ─────────────────────────────────────
function renderStats(stats) {
  const tbody = document.getElementById('statsBody');
  tbody.innerHTML = TAB_ORDER.map(tab => {
    const s = stats[tab];
    if (!s || s.total === 0) return '';
    const cls = getTabClass(tab);

    // 주차별 배지 (정렬)
    const weekHtml = Object.entries(s.by_week || {})
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([w, cnt]) => `<span class="week-badge">${w} <b>${cnt}</b>건</span>`)
      .join('');

    // 월별 배지
    const monthHtml = Object.entries(s.by_month || {})
      .sort((a, b) => parseInt(a[0]) - parseInt(b[0]))
      .map(([m, cnt]) => {
        const isPrimary = m === s.primary_month;
        return `<span class="month-badge ${isPrimary ? 'month-primary' : ''}">${m} <b>${cnt}</b>건</span>`;
      }).join('');

    // 주 건수 = 주 대상 월 건수
    const displayCount = s.total_primary != null ? s.total_primary : s.total;

    // 장애율
    const faultRate = s.fault_rate != null
      ? `<span class="rate-val ${s.fault_rate >= 1 ? 'rate-bad' : 'rate-ok'}">${s.fault_rate}%</span>`
      : '<span style="color:var(--gray)">-</span>';

    return `
      <tr>
        <td><span class="tab-label ${cls}">${tab}</span></td>
        <td class="count-cell">${displayCount}</td>
        <td class="rate-cell-stats">${faultRate}</td>
        <td class="period-cell">
          <div class="period-row">${monthHtml || '-'}</div>
          <div class="period-row">${weekHtml || ''}</div>
        </td>
        <td class="top3-cell">${s.top3 || '-'}</td>
      </tr>`;
  }).join('');
}

// ─── 신규 업로드 비교 코멘트 ─────────────────────────
const CSTYLE_MAP = {
  summary_bad:'bad', summary_good:'good', summary_same:'info',
  tab_bad:'bad', tab_good:'good', tab_same:'info',
  new_bad:'new', new_good:'new', new_same:'new',
  alert:'alert', rate_info:'info', divider:'divider',
};
const CLBL_MAP = {
  summary_bad:'전체▲', summary_good:'전체▼', summary_same:'전체→',
  tab_bad:'증가▲', tab_good:'감소▼', tab_same:'유지→',
  new_bad:'신규▲', new_good:'신규▼', new_same:'신규→',
  alert:'경보⚠', rate_info:'장애율',
};

function _renderCommentList(listEl, comments) {
  if (!comments || !comments.length) {
    listEl.innerHTML = '<li class="comment-item"><span class="comment-text" style="color:var(--gray)">데이터 없음</span></li>';
    return;
  }
  listEl.innerHTML = comments.map(c => {
    const cls = CSTYLE_MAP[c.type] || 'info';
    const lbl = CLBL_MAP[c.type] || (c.tag || '');
    if (c.type === 'divider') {
      return `<li class="comment-item divider-row"><span class="comment-text">${c.text}</span></li>`;
    }
    return `<li class="comment-item">
      <span class="c-tag ${cls}">${lbl}</span>
      <span class="comment-text">${c.text}</span>
    </li>`;
  }).join('');
}

function renderUploadComparison(comments, weekComments) {
  const card = document.getElementById('uploadCompareCard');
  const hasMonth = comments && comments.length > 0;
  const hasWeek  = weekComments && weekComments.length > 0;
  if (!hasMonth && !hasWeek) {
    card.style.display = 'none';
    return;
  }
  card.style.display = 'block';
  _renderCommentList(document.getElementById('uploadCommentList'), comments);
  _renderCommentList(document.getElementById('weekCommentList'), weekComments);
}

// ─── 월간 파일 섹션 ──────────────────────────────────
let monthlyFilesInfo = [];

async function setupMonthlySection() {
  const res = await fetch('/api/monthly-files');
  const data = await res.json();
  monthlyFilesInfo = data.files || [];

  const sel = document.getElementById('monthlyFileSelect');
  if (!monthlyFilesInfo.length) {
    sel.innerHTML = '<option value="">파일 없음 (monthly_files 폴더 확인)</option>';
    return;
  }
  sel.innerHTML = monthlyFilesInfo.map(f =>
    `<option value="${f.name}">${f.confirmed ? '🔒 ' : ''}${f.name}${f.confirmed ? ' [확정]' : ''}</option>`
  ).join('');

  // 데이터에서 감지된 월에 맞는 파일 자동 선택
  if (detectedWeeks.length > 0) {
    const week = detectedWeeks[0];
    const month = week.match(/^(\d+)월/)?.[1];
    if (month) {
      const padded = month.padStart(2, '0');
      const match = monthlyFilesInfo.find(f => f.name.includes(`${padded}월`));
      if (match) sel.value = match.name;
    }
  }

  // 주차 목록
  const wSel = document.getElementById('weekSelect');
  wSel.innerHTML = detectedWeeks.map(w => `<option value="${w}">${w}</option>`).join('');

  // 다운로드 버튼 숨기기
  document.getElementById('btnDownloadMonthly').style.display = 'none';

  // 초기 확정 상태 반영
  onFileSelectChange();
}

function onFileSelectChange() {
  const selVal = document.getElementById('monthlyFileSelect').value;
  const info = monthlyFilesInfo.find(f => f.name === selVal);
  const isConfirmed = info?.confirmed ?? false;
  const btn = document.getElementById('btnInsert');
  const notice = document.getElementById('confirmedNotice');
  btn.disabled = isConfirmed;
  if (notice) notice.style.display = isConfirmed ? 'block' : 'none';
}

function onModeChange() {
  const mode = document.getElementById('modeSelect').value;
  document.getElementById('weekGroup').style.display = mode === 'weekly' ? 'block' : 'none';
  document.getElementById('modeDesc').textContent = MODE_DESC[mode] || '';
}
// 초기 설명
document.getElementById('modeDesc').textContent = MODE_DESC['daily'];

// ─── 월간 파일 삽입 ──────────────────────────────────
document.getElementById('btnInsert').addEventListener('click', async () => {
  if (!resultFilename) return showToast('먼저 처리를 실행해주세요.', 'error');
  const monthlyFilename = document.getElementById('monthlyFileSelect').value;
  if (!monthlyFilename) return showToast('대상 파일을 선택해주세요.', 'error');
  const mode = document.getElementById('modeSelect').value;
  const weekLabel = document.getElementById('weekSelect').value;

  if (mode === 'weekly' && !weekLabel) return showToast('주차를 선택해주세요.', 'error');

  // ── 월 일치 검증 ─────────────────────────────────────
  // 파일명에서 대상 월 추출 (예: "2026년 03월..." → 3)
  const fileMonthMatch = monthlyFilename.match(/(\d{1,2})월/);
  const fileMonth = fileMonthMatch ? parseInt(fileMonthMatch[1], 10) : null;
  if (fileMonth && Object.keys(statsData).length > 0) {
    // 처리된 데이터의 주 대상 월 집계
    const monthCount = {};
    for (const tab in statsData) {
      const pm = statsData[tab].primary_month; // e.g. "3월"
      if (pm) {
        const n = parseInt(pm);
        monthCount[n] = (monthCount[n] || 0) + 1;
      }
    }
    const dataMonth = Object.entries(monthCount).sort((a, b) => b[1] - a[1])[0]?.[0];
    if (dataMonth && parseInt(dataMonth) !== fileMonth) {
      showToast(
        `선택한 파일은 ${fileMonth}월 파일이지만, 처리된 데이터는 ${dataMonth}월 데이터입니다.\n해당 월 데이터가 아닙니다.`,
        'error'
      );
      return;
    }
  }

  const modeLabel = { daily: '일별 추가', weekly: '주차 확정', monthly: '월 마감' }[mode];
  if (mode === 'monthly') {
    if (!confirm(`[월 마감] ${monthlyFilename}의 전체 데이터를 새 데이터로 교체합니다.\n계속하시겠습니까?`)) return;
  } else if (mode === 'weekly') {
    if (!confirm(`[주차 확정] ${monthlyFilename}의 "${weekLabel}" 데이터를 교체합니다.\n계속하시겠습니까?`)) return;
  }

  showLoading(true, `${modeLabel} 중...`);

  try {
    const res = await fetch('/api/insert-monthly', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        result_key: resultFilename,
        monthly_filename: monthlyFilename,
        mode,
        week_label: weekLabel,
      }),
    });
    let data;
    try {
      data = await res.json();
    } catch (_) {
      showToast('서버 메모리 부족으로 처리 실패. 잠시 후 다시 시도하거나 로컬에서 실행하세요.', 'error');
      return;
    }

    if (!data.ok) {
      showToast(data.error || '삽입 오류', 'error');
      return;
    }

    renderInsertReport(data.report || {});
    document.getElementById('insertReport').style.display = 'block';
    document.getElementById('btnDownloadMonthly').style.display = 'inline-flex';
    document.getElementById('btnDownloadMonthly').dataset.filename = monthlyFilename;

    const totalAdded = Object.values(data.report)
      .filter(r => !r.skipped)
      .reduce((s, r) => s + (r.added || 0), 0);
    showToast(`삽입 완료! ${totalAdded}건 추가`, 'success');
    document.getElementById('insertReport').scrollIntoView({ behavior: 'smooth' });

  } catch (e) {
    showToast('서버 오류: ' + e.message, 'error');
  } finally {
    showLoading(false);
  }
});

function renderInsertReport(report) {
  const tbody = document.getElementById('reportBody');
  tbody.innerHTML = Object.entries(report).map(([tab, r]) => {
    if (r.skipped) {
      return `<tr class="row-skipped">
        <td>${tab}</td>
        <td colspan="4" style="color:var(--gray);font-style:italic">건너뜀: ${r.reason || ''}</td>
      </tr>`;
    }
    const addedClass = r.added > 0 ? 'added-positive' : '';
    return `<tr>
      <td>${tab}</td>
      <td style="color:var(--gray);font-size:11px">${r.sheet}</td>
      <td>${r.before}</td>
      <td class="${addedClass}">+${r.added}</td>
      <td><b>${r.total}</b></td>
    </tr>`;
  }).join('');
}

document.getElementById('btnDownloadMonthly').addEventListener('click', () => {
  const fn = document.getElementById('btnDownloadMonthly').dataset.filename;
  if (!fn) return;
  window.location.href = '/api/download-monthly/' + encodeURIComponent(fn);
});

// ─── 유틸 ────────────────────────────────────────
function showLoading(on, msg = '처리 중입니다...') {
  document.getElementById('loadingMsg').textContent = msg;
  document.getElementById('loadingOverlay').style.display = on ? 'flex' : 'none';
}

let toastTimer;
function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast' + (type ? ' ' + type : '') + ' show';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
}
