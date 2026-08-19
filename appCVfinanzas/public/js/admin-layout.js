(function () {
  const sidebar = document.getElementById('sidebar');
  const sidebarOverlay = document.getElementById('sidebarOverlay');
  const sidebarToggle = document.getElementById('sidebarToggle');
  const sidebarClose = document.getElementById('sidebarClose');
  const userTargets = document.querySelectorAll('[data-user-display]');
  const logoutButtons = document.querySelectorAll('[data-logout]');
  const navLinks = document.querySelectorAll('[data-nav-link]');
  const themeButtons = document.querySelectorAll('[data-theme-toggle]');

  function getInitialTheme() {
    try {
      const savedTheme = localStorage.getItem('cvf-theme');
      if (savedTheme === 'dark' || savedTheme === 'light') {
        return savedTheme;
      }
    } catch (error) {}

    return document.documentElement.classList.contains('dark') ? 'dark' : 'light';
  }

  function updateThemeButtons(isDark) {
    themeButtons.forEach((button) => {
      button.setAttribute('aria-pressed', String(isDark));
      button.querySelectorAll('[data-theme-label]').forEach((label) => {
        label.textContent = isDark ? 'Tema claro' : 'Tema oscuro';
      });
      button.querySelectorAll('[data-theme-icon="dark"]').forEach((icon) => {
        icon.classList.toggle('hidden', isDark);
      });
      button.querySelectorAll('[data-theme-icon="light"]').forEach((icon) => {
        icon.classList.toggle('hidden', !isDark);
      });
    });
  }

  function applyTheme(theme) {
    const isDark = theme === 'dark';
    document.documentElement.classList.toggle('dark', isDark);

    try {
      localStorage.setItem('cvf-theme', theme);
    } catch (error) {}

    updateThemeButtons(isDark);
  }

  function bindThemeToggle() {
    applyTheme(getInitialTheme());

    themeButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const nextTheme = document.documentElement.classList.contains('dark') ? 'light' : 'dark';
        applyTheme(nextTheme);
      });
    });
  }

  function openSidebar() {
    if (!sidebar || !sidebarOverlay) return;
    sidebar.classList.remove('-translate-x-full');
    sidebarOverlay.classList.remove('hidden');
  }

  function closeSidebar() {
    if (!sidebar || !sidebarOverlay) return;
    sidebar.classList.add('-translate-x-full');
    sidebarOverlay.classList.add('hidden');
  }

  function setActiveNav() {
    const currentPath = window.location.pathname;

    navLinks.forEach((link) => {
      const linkPath = new URL(link.getAttribute('href'), window.location.origin).pathname;
      const cleanPath = currentPath.replace(/\.html$/, '');
      const cleanLinkPath = linkPath.replace(/\.html$/, '');
      const isActive = cleanPath === cleanLinkPath;

      link.classList.toggle('bg-blue-50', isActive);
      link.classList.toggle('text-blue-700', isActive);
      link.classList.toggle('border-blue-600', isActive);
      link.classList.toggle('text-gray-600', !isActive);
      link.classList.toggle('border-transparent', !isActive);
    });
  }

  async function loadUser() {
    if (!userTargets.length) return;

    try {
      const response = await fetch('/auth/me');

      if (response.status === 401) {
        window.location.href = '/login';
        return;
      }

      if (!response.ok) return;

      const data = await response.json();
      const displayName = data.user?.usuario || 'Usuario';

      userTargets.forEach((target) => {
        target.textContent = displayName;
      });
    } catch (error) {
      console.warn('No se pudo cargar la sesion activa.');
    }
  }

  function bindLogout() {
    logoutButtons.forEach((button) => {
      button.addEventListener('click', async () => {
        await fetch('/auth/logout', { method: 'POST' }).catch(() => {});
        window.location.href = '/login';
      });
    });
  }

  sidebarToggle?.addEventListener('click', openSidebar);
  sidebarClose?.addEventListener('click', closeSidebar);
  sidebarOverlay?.addEventListener('click', closeSidebar);
  bindThemeToggle();
  setActiveNav();
  bindLogout();
  loadUser();
}());
