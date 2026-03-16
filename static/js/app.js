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

// 탭명 → 월간 파일 숨김 시트명 매핑
const TAB_TO_RAWSHEET = {
  "서울 B710":           "B710 로우데이터",
  "서울 B800":           "B800로우데이터",
  "서울 B700":           "B700로우데이터",
  "공항 B620":           "공항 B620로우데이터",
  "대전 B650":           "대전 B650 로우데이터",
  "세종 B500":           "세종B500로우데이터",
  "제주 B400":           "제주 B400 로우데이터",
  "포항 B800":           "포항B800로우데이터",
  "상주,영주,예천 B400": "상주.영주.예천 로우데이터",
  "안동 B520D":          "안동B520D 로우데이터",
  "김해 B600":           "김해B600 로우데이터",
};

// ─── SheetJS 기반 클라이언트 사이드 삽입 헬퍼 ───────
function _jsColMap(sheetHeaders, procCols) {
  const nameToPos = {};
  sheetHeaders.forEach((h, i) => {
    if (h == null) return;
    const s = String(h).trim();
    if (!nameToPos[s]) nameToPos[s] = [];
    nameToPos[s].push(i);
  });
  const colMap = {};
  const used = new Set();
  for (const col of procCols) {
    const cs = String(col).trim();
    if (cs === '날짜') {
      const avail = (nameToPos['장애접수일시'] || []).filter(p => !used.has(p));
      if (avail.length >= 2) { colMap[col] = avail[1]; used.add(avail[1]); }
      else if (avail.length === 1) { colMap[col] = avail[0]; used.add(avail[0]); }
    } else if (cs === 'cits') {
      for (const cn of ['CITS 설치일', 'CITS설치일']) {
        const avail = (nameToPos[cn] || []).filter(p => !used.has(p));
        if (avail.length > 0) { colMap[col] = avail[0]; used.add(avail[0]); break; }
      }
    } else {
      const avail = (nameToPos[cs] || []).filter(p => !used.has(p));
      if (avail.length > 0) { colMap[col] = avail[0]; used.add(avail[0]); }
    }
  }
  return colMap;
}

function _jsFindColPos(headers, names) {
  for (const name of names) {
    const idx = headers.findIndex(h => h != null && String(h).trim() === name);
    if (idx >= 0) return idx;
  }
  return null;
}

function _jsFindDatePos(headers) {
  const pos = _jsFindColPos(headers, ['장애접수일']);
  if (pos != null) return pos;
  const occ = headers.reduce((a, h, i) => { if (String(h || '') === '장애접수일시') a.push(i); return a; }, []);
  return occ.length >= 2 ? occ[1] : null;
}

async function _doClientInsert(monthlyFilename, mode, weekLabel) {
  // 1. 처리된 탭 데이터 JSON 수신
  const dataRes = await fetch(`/api/result-data/${encodeURIComponent(resultFilename)}`);
  if (!dataRes.ok) {
    const err = await dataRes.json().catch(() => ({}));
    throw new Error(err.error || '처리된 데이터를 불러올 수 없습니다. 다시 처리를 실행해주세요.');
  }
  const { tabs: tabsData } = await dataRes.json();

  // 2. 월간 xlsx 파일 ArrayBuffer 수신
  const xlsxRes = await fetch(`/api/download-monthly/${encodeURIComponent(monthlyFilename)}`);
  if (!xlsxRes.ok) throw new Error('월간 파일을 불러올 수 없습니다.');
  const arrayBuffer = await xlsxRes.arrayBuffer();

  // 3. SheetJS로 워크북 읽기
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: true });
  const report = {};

  for (const [tabName, tabData] of Object.entries(tabsData)) {
    const sheetName = TAB_TO_RAWSHEET[tabName];
    if (!sheetName || !wb.SheetNames.includes(sheetName)) {
      report[tabName] = { skipped: true, reason: '시트 없음' };
      continue;
    }

    const ws = wb.Sheets[sheetName];
    const wsArray = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    let sheetHeaders = wsArray.length > 0 ? wsArray[0].map(h => h != null ? String(h) : null) : [];
    let existingRows = wsArray.length > 1 ? wsArray.slice(1).filter(r => r.some(v => v != null)) : [];

    if (!sheetHeaders.length) {
      sheetHeaders = tabData.columns;
      existingRows = [];
    }

    const numCols = sheetHeaders.length;
    const colMap = _jsColMap(sheetHeaders, tabData.columns);

    // 새 행을 시트 컬럼 순서에 맞게 변환
    const newRows = tabData.rows.map(row =>  {
      const r = new Array(numCols).fill(null);
      tabData.columns.forEach((col, ci) => {
        const pos = colMap[col];
        if (pos != null && pos < numCols) r[pos] = row[ci] ?? null;
      });
      return r;
    });

    const idPos   = _jsFindColPos(sheetHeaders, ['접수번호', 'No']);
    const weekPos = _jsFindColPos(sheetHeaders, ['주차']);
    const datePos = _jsFindDatePos(sheetHeaders);
    const before  = existingRows.length;

    let merged;
    if (mode === 'monthly') {
      merged = newRows;
    } else if (mode === 'weekly' && weekLabel) {
      const keep = weekPos != null
        ? existingRows.filter(r => String(r[weekPos] ?? '') !== String(weekLabel))
        : existingRows;
      merged = [...keep, ...newRows];
    } else {
      if (idPos != null) {
        const existingIds = new Set(existingRows.map(r => String(r[idPos] ?? '')));
        const toAdd = newRows.filter(r => !existingIds.has(String(r[idPos] ?? '')));
        merged = [...existingRows, ...toAdd];
      } else {
        merged = [...existingRows, ...newRows];
      }
    }

    // 날짜 기준 정렬
    if (datePos != null && merged.length > 0) {
      merged.sort((a, b) => {
        const da = a[datePos] != null ? String(a[datePos]) : '\uFFFF';
        const db = b[datePos] != null ? String(b[datePos]) : '\uFFFF';
        return da < db ? -1 : da > db ? 1 : 0;
      });
    }

    // 시트 업데이트
    const newWs = XLSX.utils.aoa_to_sheet([sheetHeaders, ...merged]);
    if (ws['!cols']) newWs['!cols'] = ws['!cols'];
    wb.Sheets[sheetName] = newWs;

    report[tabName] = {
      sheet: sheetName,
      before,
      added: merged.length - before,
      total: merged.length,
    };
  }

  // 4. 수정된 xlsx 다운로드
  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = monthlyFilename; a.click();
  URL.revokeObjectURL(url);

  return report;
}

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
const _btnSettings = document.getElementById('btnSettings');
if (_btnSettings) _btnSettings.addEventListener('click', () => {
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

// 확정 데이터 가져오기
document.getElementById('importConfirmedInput').addEventListener('change', function() {
  const file = this.files[0];
  if (!file) return;
  const status = document.getElementById('importConfirmedStatus');
  status.textContent = '업로드 중...';
  const fd = new FormData();
  fd.append('file', file);
  fetch('/api/import-confirmed', { method: 'POST', body: fd })
    .then(r => r.json())
    .then(d => {
      if (d.ok) {
        status.style.color = 'var(--green)';
        status.textContent = `완료! ${d.imported}개 항목 적용됨`;
      } else {
        status.style.color = 'var(--primary)';
        status.textContent = '오류: ' + d.error;
      }
    })
    .catch(() => { status.style.color='var(--primary)'; status.textContent='업로드 실패'; });
  this.value = '';
});
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

// ─── 회원 관리 ───────────────────────────────────────
let _foundUser = null;

document.getElementById('btnUserSearch').addEventListener('click', searchUser);
document.getElementById('userSearchInput').addEventListener('keydown', e => {
  if (e.key === 'Enter') searchUser();
});

function searchUser() {
  const q = document.getElementById('userSearchInput').value.trim();
  const result = document.getElementById('userSearchResult');
  const errEl = document.getElementById('userSearchError');
  result.style.display = 'none';
  errEl.style.display = 'none';
  _foundUser = null;
  if (!q) return;
  fetch('/api/admin/search-user?q=' + encodeURIComponent(q))
    .then(r => r.json())
    .then(d => {
      if (!d.ok) { errEl.textContent = d.error; errEl.style.display = 'block'; return; }
      _foundUser = d;
      document.getElementById('userResultName').textContent = d.name || d.username;
      document.getElementById('userResultId').textContent = '(' + d.username + ')';
      const badge = document.getElementById('userResultBadge');
      if (d.is_admin) {
        badge.textContent = '관리자';
        badge.style.cssText = 'margin-left:6px;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:#FFF0F4;color:var(--primary);border:1px solid #F5C0CF';
      } else {
        badge.textContent = '일반';
        badge.style.cssText = 'margin-left:6px;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:#e5e7eb;color:var(--gray-dark);border:1px solid #d1d5db';
      }
      document.getElementById('btnGrantAdmin').style.display = d.is_admin ? 'none' : 'inline-block';
      document.getElementById('btnRevokeAdmin').style.display = d.is_admin ? 'inline-block' : 'none';
      result.style.display = 'block';
    })
    .catch(() => { errEl.textContent = '검색 오류'; errEl.style.display = 'block'; });
}

function setAdminRole(isAdmin) {
  if (!_foundUser) return;
  fetch('/api/admin/set-admin', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: _foundUser.username, is_admin: isAdmin }),
  })
    .then(r => r.json())
    .then(d => {
      if (!d.ok) { showToast(d.error, 'error'); return; }
      showToast(`${_foundUser.name || _foundUser.username} → ${isAdmin ? '관리자' : '일반 계정'} 변경 완료`, 'success');
      searchUser();
    })
    .catch(() => showToast('변경 실패', 'error'));
}

document.getElementById('btnGrantAdmin').addEventListener('click', () => setAdminRole(true));
document.getElementById('btnRevokeAdmin').addEventListener('click', () => setAdminRole(false));

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
    renderUploadComparison(data.comparison || {});
    if (window.IS_ADMIN) setupMonthlySection();

    document.getElementById('resultCard').style.display = 'block';
    document.getElementById('previewCard').style.display = 'block';
    document.getElementById('statsCard').style.display = 'block';
    if (window.IS_ADMIN) {
      document.getElementById('monthlyCard').style.display = 'block';
      document.getElementById('insertReport').style.display = 'none';
    }

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

// ─── 업로드 데이터 비교 (새 UI) ──────────────────────
function _tabLabelCls(tab) {
  if (tab === '전체') return 'total';
  if (tab.startsWith('서울')) return 'seoul';
  if (tab.startsWith('공항')) return 'airport';
  return 'regional';
}

function _diffCell(diff, pct) {
  if (diff === 0) return `<span class="cmp-diff-same">→ 동일</span>`;
  const sign = diff > 0 ? '▲' : '▼';
  const cls  = diff > 0 ? 'cmp-diff-up' : 'cmp-diff-down';
  const pctStr = pct != null ? ` (${Math.abs(pct)}%)` : '';
  return `<span class="${cls}">${sign}${Math.abs(diff)}건${pctStr}</span>`;
}

function renderUploadComparison(comparison) {
  const card = document.getElementById('uploadCompareCard');
  const cmps = (comparison && comparison.comparisons) || [];
  if (!cmps.length) { card.style.display = 'none'; return; }

  card.style.display = 'block';
  const tabBtns = document.getElementById('cmpTabBtns');
  const panels  = document.getElementById('cmpPanels');

  // 탭 버튼
  const TYPE_ICON = { day: '☀', week: '📅', month: '📊' };
  tabBtns.innerHTML = cmps.map((c, i) =>
    `<button class="cmp-tab-btn${i === 0 ? ' active' : ''}" onclick="showCmpTab(${i})">
      ${TYPE_ICON[c.type] || ''} ${c.label}
      <span class="cmp-period">${c.prev_label} → ${c.cur_label}</span>
    </button>`
  ).join('');

  // 패널
  panels.innerHTML = cmps.map((c, i) => {
    const thPrev = `<th>${c.prev_label}</th>`;
    const thCur  = `<th>${c.cur_label}</th>`;
    const rows = c.rows.map(r => {
      const cls = _tabLabelCls(r.tab);
      // 오류유형 칩
      let faultHtml = '';
      if (!r.is_total && r.fault_changes && r.fault_changes.length) {
        const chips = r.fault_changes.map(f => {
          const chipCls = f.diff > 0 ? 'up' : 'down';
          const arrow = f.diff > 0 ? '▲' : '▼';
          return `<span class="fault-chip ${chipCls}">${f.name} ${arrow}${Math.abs(f.diff)}</span>`;
        }).join('');
        faultHtml = `<div class="fault-chips">${chips}</div>`;
      }
      const tabCell = r.is_total
        ? `<td><span class="cmp-tab-label total">전체</span></td>`
        : `<td><span class="cmp-tab-label ${cls}">${r.tab}</span>${faultHtml}</td>`;
      return `<tr class="${r.is_total ? 'cmp-total' : ''}">
        ${tabCell}
        <td class="cmp-num">${r.prev}</td>
        <td class="cmp-num">${r.cur}</td>
        <td>${_diffCell(r.diff, r.pct)}</td>
      </tr>`;
    }).join('');
    return `<div class="cmp-panel${i === 0 ? ' active' : ''}">
      <table class="cmp-table">
        <thead><tr><th>단말</th>${thPrev}${thCur}<th>변화</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
  }).join('');
}

function showCmpTab(idx) {
  document.querySelectorAll('.cmp-tab-btn').forEach((b, i) => b.classList.toggle('active', i === idx));
  document.querySelectorAll('.cmp-panel').forEach((p, i) => p.classList.toggle('active', i === idx));
}

// ─── 현황 요약 (홈 빈 공간) ──────────────────────────
async function loadQuickStats() {
  try {
    const [statsRes, cwRes] = await Promise.all([
      fetch('/api/analysis-data'),
      fetch('/api/confirmed-weeks'),
    ]);
    const statsData = await statsRes.json();
    const cw = await cwRes.json();

    if (!statsData.ok) return;

    const allStats = statsData.all_stats || {};
    const sortedKeys = Object.keys(allStats).sort();
    if (!sortedKeys.length) return;

    // 최신 확정 월 찾기
    const confirmedKeys = sortedKeys.filter(k => allStats[k].confirmed);
    const latestKey = confirmedKeys.length ? confirmedKeys[confirmedKeys.length - 1] : sortedKeys[sortedKeys.length - 1];
    const latest = allStats[latestKey];

    document.getElementById('quickStatsTitle').textContent =
      `${latest.label} 확정 현황`;

    // 탭별 카드
    const TAB_CLS = t => t.startsWith('서울') ? 'seoul' : t.startsWith('공항') ? 'airport' : 'regional';
    const tabCards = document.getElementById('quickTabCards');
    tabCards.innerHTML = ['서울 B800','서울 B700','서울 B710','공항 B620',
      '대전 B650','세종 B500','제주 B400','포항 B800',
      '상주,영주,예천 B400','안동 B520D','김해 B600'].map(tab => {
      const d = (latest.tabs || {})[tab] || {};
      const cnt = d.count ?? '-';
      const rate = d.fault_rate != null ? d.fault_rate : null;
      const rateCls = rate == null ? 'none' : rate >= 1 ? 'bad' : 'ok';
      const rateTxt = rate != null ? rate + '%' : '-';
      return `<div class="quick-tab-card">
        <div class="quick-tab-name">${tab}</div>
        <div class="quick-tab-count">${cnt}</div>
        <div class="quick-tab-rate ${rateCls}">장애율 ${rateTxt}</div>
      </div>`;
    }).join('');

    // 이번 달 주차 확정 현황 + 다운로드 버튼
    // 현재 달 = 가장 최신 월간 파일 기준
    const monthlyFiles = await (await fetch('/api/monthly-files')).json();
    const latestFile = (monthlyFiles.files || []).filter(f => !f.confirmed).slice(-1)[0]
      || (monthlyFiles.files || []).slice(-1)[0];
    if (latestFile) {
      const fn = latestFile.name;
      const confirmed = cw[fn] || [];
      // 주차 목록: 이미 확정된 것 + 예상 주차 (1~5주)
      const monthMatch = fn.match(/(\d{1,2})월/);
      const mn = monthMatch ? parseInt(monthMatch[1]) : null;
      if (mn && confirmed.length > 0) {
        const maxWeek = 5;
        const badges = Array.from({length: maxWeek}, (_, i) => {
          const label = `${mn}월${i+1}주`;
          const done = confirmed.includes(label);
          return `<span class="week-confirm-badge ${done ? 'done' : 'pending'}">${done ? '✓' : '·'} ${label}</span>`;
        }).join('');
        const weekRow = document.getElementById('quickWeekRow');
        document.getElementById('quickWeekBadges').innerHTML = badges;
        weekRow.style.display = 'flex';
      }
      // 최신 월간 파일 다운로드 버튼
      const dlRow = document.getElementById('quickDownloadRow');
      const dlBtn = document.getElementById('btnQuickDownloadMonthly');
      if (dlRow && dlBtn) {
        dlBtn.textContent = `⬇ ${fn} 다운로드`;
        dlBtn.onclick = () => window.location.href = '/api/download-monthly/' + encodeURIComponent(fn);
        dlRow.style.display = 'block';
      }
    }

  } catch (e) {
    document.getElementById('quickStatsTitle').textContent = '현황 불러오기 실패';
  }
}

// 페이지 로드 시 현황 표시
loadQuickStats();

// ─── 월간 파일 섹션 ──────────────────────────────────
let monthlyFilesInfo = [];

async function setupMonthlySection() {
  const [res1, _cw] = await Promise.all([
    fetch('/api/monthly-files'),
    loadConfirmedWeeks(),
  ]);
  const data = await res1.json();
  monthlyFilesInfo = data.files || [];

  const sel = document.getElementById('monthlyFileSelect');
  // 확정(월마감)된 파일은 드롭다운에서 제외
  const activeFiles = monthlyFilesInfo.filter(f => !f.confirmed);
  if (!activeFiles.length) {
    sel.innerHTML = '<option value="">진행 중인 파일 없음</option>';
    onFileSelectChange();
    return;
  }
  sel.innerHTML = activeFiles.map(f =>
    `<option value="${f.name}">${f.name}</option>`
  ).join('');

  // 데이터에서 감지된 월에 맞는 파일 자동 선택
  if (detectedWeeks.length > 0) {
    const week = detectedWeeks[0];
    const month = week.match(/^(\d+)월/)?.[1];
    if (month) {
      const padded = month.padStart(2, '0');
      const match = activeFiles.find(f => f.name.includes(`${padded}월`));
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

// ─── 확정 주차 캐시 ─────────────────────────────────
let confirmedWeeksCache = {};   // { filename: ["3월1주차", ...] }
let insertedKeys = new Set();   // "filename::weekLabel" — 이번 세션에서 삽입 완료된 키

async function loadConfirmedWeeks() {
  try {
    const res = await fetch('/api/confirmed-weeks');
    confirmedWeeksCache = await res.json();
  } catch (_) {}
}

function isWeekConfirmed(filename, weekLabel) {
  return (confirmedWeeksCache[filename] || []).includes(weekLabel);
}

function updateWeekActionUI() {
  const filename = document.getElementById('monthlyFileSelect').value;
  const weekLabel = document.getElementById('weekSelect').value;
  const mode = document.getElementById('modeSelect').value;
  const row = document.getElementById('weekActionRow');

  if (mode !== 'weekly' || !weekLabel || !filename) {
    row.style.display = 'none';
    return;
  }
  row.style.display = 'flex';

  const confirmed = isWeekConfirmed(filename, weekLabel);
  const badge = document.getElementById('weekStatusBadge');
  const btnConfirm = document.getElementById('btnConfirmWeek');
  const btnUnconfirm = document.getElementById('btnUnconfirmWeek');
  const btnDelete = document.getElementById('btnDeleteWeek');
  const btnInsert = document.getElementById('btnInsert');

  if (confirmed) {
    badge.textContent = '🔒 확정됨';
    badge.className = 'week-badge week-badge-confirmed';
    btnConfirm.style.display = 'none';
    btnUnconfirm.style.display = 'inline-flex';
    btnDelete.style.display = 'none';
    btnInsert.disabled = true;
  } else {
    badge.textContent = '🟡 미확정';
    badge.className = 'week-badge week-badge-pending';
    btnConfirm.style.display = 'inline-flex';
    btnUnconfirm.style.display = 'none';
    btnDelete.style.display = 'inline-flex';
    // 파일 자체가 확정인지 다시 체크
    const info = monthlyFilesInfo.find(f => f.name === filename);
    btnInsert.disabled = info?.confirmed ?? false;
  }
}

function onFileSelectChange() {
  const selVal = document.getElementById('monthlyFileSelect').value;
  const info = monthlyFilesInfo.find(f => f.name === selVal);
  const isConfirmed = info?.confirmed ?? false;
  const btn = document.getElementById('btnInsert');
  const notice = document.getElementById('confirmedNotice');
  btn.disabled = isConfirmed;
  if (notice) notice.style.display = isConfirmed ? 'block' : 'none';
  updateWeekActionUI();
}

function onModeChange() {
  const mode = document.getElementById('modeSelect').value;
  document.getElementById('weekGroup').style.display = mode === 'weekly' ? 'block' : 'none';
  document.getElementById('modeDesc').textContent = MODE_DESC[mode] || '';
  updateWeekActionUI();
}

function onWeekSelectChange() {
  updateWeekActionUI();
}

// ─── 주차 확정 버튼 ──────────────────────────────────
const _btnConfirmWeek = document.getElementById('btnConfirmWeek');
if (_btnConfirmWeek) _btnConfirmWeek.addEventListener('click', async () => {
  const filename = document.getElementById('monthlyFileSelect').value;
  const weekLabel = document.getElementById('weekSelect').value;
  if (!filename || !weekLabel) return;
  // 이번 세션에서 해당 파일+주차 삽입 여부 확인
  const insertKey = `${filename}::${weekLabel}`;
  if (!insertedKeys.has(insertKey)) {
    showToast(`먼저 [${weekLabel}] 데이터를 월간 파일에 삽입한 후 확정하세요.`, 'error');
    return;
  }
  const pw = prompt(`[${weekLabel}] 확정하려면 비밀번호를 입력하세요.\n확정 후에는 삽입·삭제가 불가하며 취소할 수 없습니다.`);
  if (pw === null) return;
  if (pw !== '951009') { showToast('비밀번호가 올바르지 않습니다.', 'error'); return; }
  try {
    const res = await fetch('/api/confirm-week', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ monthly_filename: filename, week_label: weekLabel }),
    });
    const data = await res.json();
    if (data.ok) {
      confirmedWeeksCache[filename] = confirmedWeeksCache[filename] || [];
      if (!confirmedWeeksCache[filename].includes(weekLabel)) confirmedWeeksCache[filename].push(weekLabel);
      updateWeekActionUI();
      showToast(`${weekLabel} 확정 완료`, 'success');
      // 마지막 업데이트 배지 갱신
      try { const lu = await fetch('/api/last-update'); if (lu.ok) renderLastUpdate(await lu.json()); } catch (_) {}
    } else {
      showToast(data.error || '오류', 'error');
    }
  } catch (e) {
    showToast('서버 오류: ' + e.message, 'error');
  }
});

// ─── 확정 취소 버튼 ─────────────────────────────────
const _btnUnconfirmWeek = document.getElementById('btnUnconfirmWeek');
if (_btnUnconfirmWeek) _btnUnconfirmWeek.addEventListener('click', async () => {
  const filename = document.getElementById('monthlyFileSelect').value;
  const weekLabel = document.getElementById('weekSelect').value;
  if (!filename || !weekLabel) return;
  const pw = prompt(`[${weekLabel}] 확정을 취소합니다.\n비밀번호를 입력하세요.`);
  if (pw === null) return;
  if (pw !== '951009') { showToast('비밀번호가 올바르지 않습니다.', 'error'); return; }
  try {
    const res = await fetch('/api/unconfirm-week', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ monthly_filename: filename, week_label: weekLabel }),
    });
    const data = await res.json();
    if (data.ok) {
      const arr = confirmedWeeksCache[filename] || [];
      const idx = arr.indexOf(weekLabel);
      if (idx !== -1) arr.splice(idx, 1);
      updateWeekActionUI();
      showToast(`${weekLabel} 확정 취소됨`, 'success');
    } else {
      showToast(data.error || '오류', 'error');
    }
  } catch (e) {
    showToast('서버 오류: ' + e.message, 'error');
  }
});

// ─── 주차 데이터 삭제 버튼 ───────────────────────────
const _btnDeleteWeek = document.getElementById('btnDeleteWeek');
if (_btnDeleteWeek) _btnDeleteWeek.addEventListener('click', async () => {
  const filename = document.getElementById('monthlyFileSelect').value;
  const weekLabel = document.getElementById('weekSelect').value;
  if (!filename || !weekLabel) return;
  if (!confirm(`[${weekLabel}] 주차 데이터를 삭제합니다.\n이 작업은 되돌릴 수 없습니다. 계속하시겠습니까?`)) return;
  showLoading(true, '주차 데이터 삭제 중...');
  try {
    const res = await fetch('/api/delete-weekly', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ monthly_filename: filename, week_label: weekLabel }),
    });
    let data;
    try { data = await res.json(); } catch (_) {
      showToast('서버 오류: 응답을 받지 못했습니다.', 'error');
      return;
    }
    if (data.ok) {
      const total = Object.values(data.report).reduce((s, r) => s + (r.deleted || 0), 0);
      showToast(`${weekLabel} 데이터 ${total}건 삭제 완료`, 'success');
    } else {
      showToast(data.error || '삭제 오류', 'error');
    }
  } catch (e) {
    showToast('서버 오류: ' + e.message, 'error');
  } finally {
    showLoading(false);
  }
});

// 초기 설명
const _modeDesc = document.getElementById('modeDesc');
if (_modeDesc) _modeDesc.textContent = MODE_DESC['daily'];

// ─── 월간 파일 삽입 ──────────────────────────────────
const _btnInsert = document.getElementById('btnInsert');
if (_btnInsert) _btnInsert.addEventListener('click', async () => {
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

  showLoading(true, `${modeLabel} 중... (파일을 처리합니다)`);

  try {
    // 서버에서 삽입 처리
    const res = await fetch('/api/insert-monthly', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ result_key: resultFilename, monthly_filename: monthlyFilename, mode, week_label: weekLabel }),
    });
    const json = await res.json();
    if (!res.ok || json.error) throw new Error(json.error || '삽입 실패');
    const report = json.report;

    // 삽입 완료 기록 (확정 시 검증용)
    insertedKeys.add(`${monthlyFilename}::${weekLabel}`);

    renderInsertReport(report);
    document.getElementById('insertReport').style.display = 'block';
    const btn = document.getElementById('btnDownloadMonthly');
    btn.style.display = 'inline-flex';
    btn.dataset.filename = monthlyFilename;

    const totalAdded = Object.values(report)
      .filter(r => !r.skipped)
      .reduce((s, r) => s + (r.added || 0), 0);
    showToast(`삽입 완료! ${totalAdded}건 추가 — 다운로드 버튼을 눌러 파일을 받으세요.`, 'success');

    // 마지막 업데이트 배지 갱신
    try {
      const lu = await fetch('/api/last-update');
      if (lu.ok) renderLastUpdate(await lu.json());
    } catch (_) {}
    document.getElementById('insertReport').scrollIntoView({ behavior: 'smooth' });

  } catch (e) {
    showToast('오류: ' + e.message, 'error');
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

const _btnDownloadMonthly = document.getElementById('btnDownloadMonthly');
if (_btnDownloadMonthly) _btnDownloadMonthly.addEventListener('click', () => {
  const fn = _btnDownloadMonthly.dataset.filename;
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

// ─── 마지막 업데이트 표시 ─────────────────────────────────
function renderLastUpdate(data) {
  const badge = document.getElementById('lastUpdateBadge');
  if (!badge || !data || !data.timestamp) return;
  const detail = data.detail ? ` <strong>${data.detail}</strong>` : '';
  badge.innerHTML = `${data.action}${detail}<br>${data.timestamp}`;
  badge.style.display = 'block';
}

// (마지막 업데이트 초기 로드는 theme.js에서 처리)
