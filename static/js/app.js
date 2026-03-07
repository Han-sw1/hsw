let selectedFiles = [];
let resultFilename = null;
let previewData = {};

const TAB_ORDER = [
  "서울 B800", "서울 B700", "서울 B710", "공항 B620",
  "대전 B650", "세종 B500", "제주 B400", "포항 B800",
  "상주,영주,예천 B400", "안동 B520D", "김해 B600"
];

function getTabClass(tab) {
  if (tab.startsWith("서울")) return "seoul";
  if (tab.startsWith("공항")) return "airport";
  return "regional";
}

// ─── 설정 모달 ───────────────────────────────────────
document.getElementById('btnSettings').addEventListener('click', () => {
  fetch('/api/config').then(r => r.json()).then(cfg => {
    document.getElementById('criteriaPath').value = cfg.criteria_path || '';
    document.getElementById('citsPath').value = cfg.cits_path || '';
    document.getElementById('settingsModal').classList.add('open');
  });
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
  }).then(r => r.json()).then(() => {
    closeSettings();
    showToast('설정이 저장되었습니다.', 'success');
  });
});

// ─── 파일 선택 ───────────────────────────────────────
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');

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
    document.getElementById('resultCard').style.display = 'none';
    document.getElementById('previewCard').style.display = 'none';
  }
}

function renderFileList(confirmed = false) {
  const list = document.getElementById('fileList');
  list.innerHTML = selectedFiles.map(f => `
    <div class="file-item">
      <span class="file-icon">📄</span>
      <span class="file-name">${f.name}</span>
      <span class="file-type-badge ${confirmed ? 'badge-confirmed' : 'badge-pending'}"
            data-name="${f.name}">${confirmed ? '✓ 확인됨' : '감지 전'}</span>
      <button class="btn-remove-file" onclick="removeFile('${f.name.replace(/'/g, "\\'")}')">✕</button>
    </div>
  `).join('');
}

// ─── 처리 ──────────────────────────────────────────
document.getElementById('btnProcess').addEventListener('click', async () => {
  if (!selectedFiles.length) return;
  showLoading(true);

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

    // 파일 배지 → "확인됨"
    renderFileList(true);

    // 요약 + 미리보기
    renderSummary(data.summary || {});
    renderTabButtons();
    document.getElementById('resultCard').style.display = 'block';
    document.getElementById('previewCard').style.display = 'block';

    // 첫 번째 데이터 있는 탭 선택
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

// ─── 다운로드 ─────────────────────────────────────
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

  // 요약 카드 활성화
  document.querySelectorAll('.summary-item').forEach((el, i) => {
    el.classList.toggle('active', TAB_ORDER[i] === tab);
  });
  // 탭 버튼 활성화
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

  const HIGHLIGHT = new Set(['날짜', 'cits', '월', '주차', '장애접수일']);
  head.innerHTML = '<tr>' + info.cols.map(c => `<th>${c}</th>`).join('') + '</tr>';
  body.innerHTML = info.rows.map(row =>
    '<tr>' + info.cols.map(c =>
      `<td class="${HIGHLIGHT.has(c) ? 'col-highlight' : ''}">${row[c] ?? ''}</td>`
    ).join('') + '</tr>'
  ).join('');
}

// ─── 유틸 ────────────────────────────────────────
function showLoading(on) {
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
