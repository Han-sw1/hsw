let selectedFile = null;
let resultFilename = null;

// ─── 설정 모달 ───
document.getElementById('btnSettings').addEventListener('click', () => {
  fetch('/api/config').then(r => r.json()).then(cfg => {
    document.getElementById('criteriaPath').value = cfg.criteria_path || '';
    document.getElementById('citsPath').value = cfg.cits_path || '';
    document.getElementById('settingsModal').classList.add('open');
  });
});
document.getElementById('closeSettings').addEventListener('click', closeSettings);
document.getElementById('cancelSettings').addEventListener('click', closeSettings);
document.getElementById('settingsModal').addEventListener('click', (e) => {
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

// ─── 파일 선택 ───
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');

fileInput.addEventListener('change', (e) => {
  if (e.target.files[0]) setFile(e.target.files[0]);
});

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const f = e.dataTransfer.files[0];
  if (f && (f.name.endsWith('.xls') || f.name.endsWith('.xlsx'))) {
    setFile(f);
  } else {
    showToast('.xls 또는 .xlsx 파일만 가능합니다.', 'error');
  }
});

function setFile(f) {
  selectedFile = f;
  document.getElementById('dropZone').style.display = 'none';
  document.getElementById('selectedFile').style.display = 'flex';
  document.getElementById('fileName').textContent = f.name;
  document.getElementById('btnProcess').disabled = false;
}

document.getElementById('removeFile').addEventListener('click', () => {
  selectedFile = null;
  fileInput.value = '';
  document.getElementById('dropZone').style.display = 'block';
  document.getElementById('selectedFile').style.display = 'none';
  document.getElementById('btnProcess').disabled = true;
  document.getElementById('resultSummary').style.display = 'none';
  document.getElementById('previewCard').style.display = 'none';
});

// ─── 처리 ───
document.getElementById('btnProcess').addEventListener('click', async () => {
  if (!selectedFile) return;

  showLoading(true);

  const formData = new FormData();
  formData.append('file', selectedFile);

  try {
    const res = await fetch('/api/process', { method: 'POST', body: formData });
    const data = await res.json();

    if (!data.ok) {
      showToast(data.error || '처리 오류', 'error');
      return;
    }

    resultFilename = data.filename;

    // 요약
    document.getElementById('statTotal').textContent = data.meta['원본_전체'];
    document.getElementById('statB700B800').textContent = data.meta['B700B800_필터'];
    document.getElementById('statReplay').textContent = data.meta['재현_필터'];
    document.getElementById('statFinal').textContent = data.meta['장애_최종'];
    document.getElementById('resultSummary').style.display = 'block';

    // 미리보기
    renderPreview(data.preview, data.preview_cols);
    document.getElementById('previewCard').style.display = 'block';

    showToast(`처리 완료! 장애 ${data.meta['장애_최종']}건`, 'success');

    document.getElementById('resultSummary').scrollIntoView({ behavior: 'smooth' });

  } catch (e) {
    showToast('서버 오류: ' + e.message, 'error');
  } finally {
    showLoading(false);
  }
});

// ─── 다운로드 ───
document.getElementById('btnDownload').addEventListener('click', () => {
  if (!resultFilename) return;
  window.location.href = '/api/download/' + encodeURIComponent(resultFilename);
});

// ─── 미리보기 테이블 ───
function renderPreview(rows, cols) {
  const highlight = new Set(['날짜', 'cits', '월', '주차']);
  const thead = document.getElementById('previewHead');
  const tbody = document.getElementById('previewBody');

  thead.innerHTML = '<tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';
  tbody.innerHTML = rows.map(row =>
    '<tr>' + cols.map(c =>
      `<td class="${highlight.has(c) ? 'col-highlight' : ''}">${row[c] ?? ''}</td>`
    ).join('') + '</tr>'
  ).join('');
}

// ─── 유틸 ───
function showLoading(on) {
  document.getElementById('loadingOverlay').style.display = on ? 'flex' : 'none';
}

let toastTimer;
function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast' + (type ? ' ' + type : '') + ' show';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3000);
}
