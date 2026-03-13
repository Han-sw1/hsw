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

  // 저장된 테마 불러오기
  const saved = localStorage.getItem(STORAGE_KEY);
  applyTheme(saved === 'dark');

  // 버튼 클릭 토글
  if (btn) {
    btn.addEventListener('click', function() {
      const isDark = document.body.classList.contains('dark');
      applyTheme(!isDark);
      localStorage.setItem(STORAGE_KEY, isDark ? 'light' : 'dark');
    });
  }
})();
