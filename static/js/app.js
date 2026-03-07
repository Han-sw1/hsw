let selectedFiles = [];
let resultFilename = null;

// 파일 타입별 배지 정보
const FILE_TYPE_INFO = {
  "서울":    { label: "서울 B700/B800", cls: "badge-seoul" },
  "B710":    { label: "서울 B710",      cls: "badge-seoul" },
  "B620":    { label: "공항 B620",      cls: "badge-airport" },
  "지역버스": { label: "지역버스",       cls: "badge-regional" },
  "unknown": { label: "미확인",         cls: "badge-unknown" },
};

// 탭 → 색상 클래스
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
  files.forEach(f => {
    if (!existing.has(f.name)) selectedFiles.push(f);
  });
  renderFileList();
  document.getElementById('btnProcess').disabled = selectedFiles.length === 0;
}

function removeFile(name) {
  selectedFiles = selectedFiles.filter(f => f.name !== name);
  renderFileList();
  document.getElementById('btnProcess').disabled = selectedFiles.length === 0;
  if (!selectedFiles.length) {
    document.getElementById('resultCard').style.display = 'none';
  }
}

function renderFileList() {
  const list = document.getElementById('fileList');
  list.innerHTML = selectedFiles.map(f => `
    <div class="file-item">
      <span class="file-icon">📄</span>
      <span class="file-name">${f.name}</span>
      <span class="file-type-badge badge-unknown" data-name="${f.name}">감지 전</span>
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

    // 파일 타입 배지 업데이트
    if (data.file_types) {
      document.querySelectorAll('.file-type-badge').forEach(badge => {
        const name = badge.dataset.name;
        const type = data.file_types[name] || 'unknown';
        const info = FILE_TYPE_INFO[type] || FILE_TYPE_INFO['unknown'];
        badge.textContent = info.label;
        badge.className = `file-type-badge ${info.cls}`;
      });
    }

    // 요약
    renderSummary(data.summary);
    document.getElementById('resultCard').style.display = 'block';

    const total = Object.values(data.summary).reduce((a, b) => a + b, 0);
    showToast(`처리 완료! 총 ${total}건 (${Object.keys(data.summary).length}개 탭)`, 'success');
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
  grid.innerHTML = Object.entries(summary).map(([tab, count]) => `
    <div class="summary-item ${getTabClass(tab)}">
      <div class="summary-tab">${tab}</div>
      <div class="summary-count">${count}</div>
      <div class="summary-label">건</div>
    </div>
  `).join('');
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
