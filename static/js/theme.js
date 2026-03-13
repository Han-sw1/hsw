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

// 사이드바 열고 닫기
(function() {
  const STORAGE_KEY = 'atmo_sidebar';
  const btn = document.getElementById('sidebarToggle');
  const backdrop = document.getElementById('sidebarBackdrop');
  const MOBILE_BP = 900;

  function isMobile() { return window.innerWidth <= MOBILE_BP; }

  function closeMobile() {
    document.body.classList.remove('sidebar-mobile-open');
  }

  // 데스크탑: 저장된 상태 복원
  if (!isMobile() && localStorage.getItem(STORAGE_KEY) === 'closed') {
    document.body.classList.add('sidebar-collapsed');
  }

  if (btn) {
    btn.addEventListener('click', function() {
      if (isMobile()) {
        document.body.classList.toggle('sidebar-mobile-open');
      } else {
        const collapsed = document.body.classList.toggle('sidebar-collapsed');
        localStorage.setItem(STORAGE_KEY, collapsed ? 'closed' : 'open');
      }
    });
  }

  // 백드롭 클릭 시 닫기
  if (backdrop) {
    backdrop.addEventListener('click', closeMobile);
  }

  // 화면 크기 변경 시 모바일 오버레이 정리
  window.addEventListener('resize', function() {
    if (!isMobile()) {
      closeMobile();
    }
  });
})();
