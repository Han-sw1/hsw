// 다크모드 토글
(function() {
  const STORAGE_KEY = 'atmo_theme';
  const btn = document.getElementById('themeToggle');

  function applyTheme(dark) {
    if (dark) {
      document.body.classList.add('dark');
      if (btn) btn.textContent = '🌙';
    } else {
      document.body.classList.remove('dark');
      if (btn) btn.textContent = '☀️';
    }
  }

  const saved = localStorage.getItem(STORAGE_KEY);
  applyTheme(saved === 'dark');

  if (btn) {
    btn.addEventListener('click', function() {
      const isDark = document.body.classList.contains('dark');
      applyTheme(!isDark);
      localStorage.setItem(STORAGE_KEY, isDark ? 'light' : 'dark');
    });
  }
})();

// 마지막 업데이트 배지
(function() {
  var badge = document.getElementById('lastUpdateBadge');
  if (!badge) return;
  fetch('/api/last-update').then(function(r) {
    if (!r.ok) return;
    return r.json();
  }).then(function(data) {
    if (!data || !data.timestamp) return;
    var detail = data.detail ? ' <strong>' + data.detail + '</strong>' : '';
    badge.innerHTML = data.action + detail + '<br>' + data.timestamp;
    badge.style.display = 'block';
  }).catch(function() {});
})();

// 사이드바 열고 닫기
(function() {
  const STORAGE_KEY = 'atmo_sidebar_v2';
  const btn = document.getElementById('sidebarToggle');
  const backdrop = document.getElementById('sidebarBackdrop');
  const MOBILE_BP = 900;

  function isMobile() { return window.innerWidth <= MOBILE_BP; }

  function initSidebar() {
    if (isMobile()) {
      // 모바일: desktop 클래스 제거, 기본 숨김 상태
      document.body.classList.remove('sidebar-collapsed');
      document.body.classList.remove('sidebar-mobile-open');
    } else {
      // 데스크탑: 저장된 상태 복원 (기본값: 닫힘)
      document.body.classList.remove('sidebar-mobile-open');
      if (localStorage.getItem(STORAGE_KEY) === 'open') {
        document.body.classList.remove('sidebar-collapsed');
      } else {
        document.body.classList.add('sidebar-collapsed');
      }
    }
  }

  initSidebar();

  if (btn) {
    btn.addEventListener('click', function() {
      if (isMobile()) {
        document.body.classList.toggle('sidebar-mobile-open');
      } else {
        const collapsed = document.body.classList.toggle('sidebar-collapsed');
        localStorage.setItem(STORAGE_KEY, collapsed ? 'closed' : 'open');
        setTimeout(function() { window.dispatchEvent(new Event('resize')); }, 260);
      }
    });
  }

  if (backdrop) {
    backdrop.addEventListener('click', function() {
      document.body.classList.remove('sidebar-mobile-open');
    });
  }

  let resizeTimer;
  window.addEventListener('resize', function() {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(initSidebar, 100);
  });
})();
