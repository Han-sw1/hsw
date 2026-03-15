const TAB_ORDER = [
  "서울 B800","서울 B700","서울 B710","공항 B620",
  "대전 B650","세종 B500","제주 B400","포항 B800",
  "상주,영주,예천 B400","안동 B520D","김해 B600"
];
function getTabClass(t){ return t.startsWith("서울")?"seoul":t.startsWith("공항")?"airport":"regional"; }

let confirmedWeeksData = [];

// ─── 초기 로드 ──────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => loadWeekly());
document.addEventListener('visibilitychange', () => { if (!document.hidden) loadWeekly(); });

function loadWeekly() {
  fetch('/api/confirmed-weeks-stats', { cache:'no-store' })
    .then(r => r.json())
    .then(data => {
      if (!data.ok || !data.weeks || data.weeks.length === 0) {
        document.getElementById('loadingBar').textContent = '확정된 주차 데이터가 없습니다.';
        return;
      }
      confirmedWeeksData = data.weeks;
      document.getElementById('loadingBar').style.display = 'none';
      document.getElementById('dashContent').style.display = 'block';

      // 확정된 주차 드롭다운
      const sel = document.getElementById('confirmedWeekSel');
      sel.innerHTML = [...data.weeks].reverse().map((w, i) => {
        const lock = w.confirmed ? ' 🔒' : '';
        return `<option value="${data.weeks.length - 1 - i}">${w.week_label}${lock}</option>`;
      }).join('');
      renderConfirmedWeekDetail();

      // 주차 비교 드롭다운
      if (data.weeks.length >= 2) {
        const optHtml = [...data.weeks].reverse().map((w, i) => {
          const lock = w.confirmed ? ' 🔒' : '';
          return `<option value="${data.weeks.length - 1 - i}">${w.week_label}${lock}</option>`;
        }).join('');
        document.getElementById('weekSelA').innerHTML = optHtml;
        document.getElementById('weekSelB').innerHTML = optHtml;
        document.getElementById('weekSelA').value = data.weeks.length - 2;
        document.getElementById('weekSelB').value = data.weeks.length - 1;
        renderWeekCompare();
      }
    })
    .catch(e => {
      document.getElementById('loadingBar').textContent = '서버 연결 오류: ' + e.message;
      document.getElementById('loadingBar').style.color = 'var(--primary)';
    });
}

// ─── 확정된 주차 현황 ─────────────────────────────────
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

  const maxCnt = Math.max(...TAB_ORDER.map(t => Math.max(wA.tab_counts[t]||0, wB.tab_counts[t]||0)), 1);

  const rows = TAB_ORDER.map(tab => {
    const cA=wA.tab_counts[tab]||0, cB=wB.tab_counts[tab]||0;
    if (cA===0 && cB===0) return '';
    const rA=wA.tab_rates?.[tab], rB=wB.tab_rates?.[tab];
    const diff=cB-cA, cls=diff>0?'diff-up':diff<0?'diff-down':'diff-same';
    const arrow=diff>0?'▲':diff<0?'▼':'→';
    const pct=cA>0?` (${diff>0?'+':''}${Math.round(diff/cA*100)}%)`:'';
    const barW=Math.round(Math.abs(diff)/maxCnt*60);
    const barColor=diff>0?'var(--primary)':diff<0?'var(--green)':'var(--gray)';
    const barHtml=diff!==0?`<span class="cmp-bar" style="width:${barW}px;background:${barColor}"></span>`:'';
    const rateCellA=rA!=null?`<span class="cmp-rate ${rA>=1?'bad':'ok'}">장애율 ${rA}%</span>`:`<span class="cmp-rate">-</span>`;
    const rateCellB=rB!=null?`<span class="cmp-rate ${rB>=1?'bad':'ok'}">장애율 ${rB}%</span>`:`<span class="cmp-rate">-</span>`;
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
      <th>단말기</th><th>${wA.week_label}</th><th>${wB.week_label}</th><th>변화량</th>
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

  // 장애유형 변화
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
    if (a===0&&b>0)       newFaults.push({ft,a,b,d:b});
    else if (b===0&&a>0)  goneFaults.push({ft,a,b,d:-a});
    else if (b>a)         increased.push({ft,a,b,d:b-a});
    else if (b<a)         decreased.push({ft,a,b,d:b-a});
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

  const noChange = !increased.length&&!decreased.length&&!newFaults.length&&!goneFaults.length;
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
