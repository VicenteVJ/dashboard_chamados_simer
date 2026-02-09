// Portal (abas + iframe)
// Regras:
// - Uma aba ativa por vez
// - Acessibilidade: role=tablist/tab, aria-selected, teclado
// - Suporte a hash: #dashboard | #ocorrencias
// - Persiste última aba no localStorage

(function () {
  const STORAGE_KEY = 'portal_active_tab_v1';

  function qs(sel, root = document) { return root.querySelector(sel); }
  function qsa(sel, root = document) { return Array.from(root.querySelectorAll(sel)); }

  const tablist = qs('[data-tabs]');
  const frame = qs('#frame');
  const tabs = qsa('[data-tab]', tablist);

  if (!tablist || !frame || tabs.length === 0) return;

  function getTabById(tabId) {
    return tabs.find(t => t.getAttribute('data-tab') === tabId) || null;
  }

  function getTabIdFromHash() {
    const h = (location.hash || '').replace('#', '').trim();
    return h || null;
  }

  function saveLastTab(tabId) {
    try { localStorage.setItem(STORAGE_KEY, tabId); } catch (_) { }
  }

  function loadLastTab() {
    try { return localStorage.getItem(STORAGE_KEY); } catch (_) { return null; }
  }

  function setActiveTab(tabId, opts = {}) {
    const { updateHash = true, focus = false } = opts;

    const btn = getTabById(tabId) || tabs[0];
    const url = btn.getAttribute('data-url');
    if (url) frame.setAttribute('src', url);

    tabs.forEach(t => {
      const isActive = t === btn;
      t.classList.toggle('active', isActive);
      t.setAttribute('aria-selected', isActive ? 'true' : 'false');
      t.setAttribute('tabindex', isActive ? '0' : '-1');
    });

    const activeId = btn.getAttribute('data-tab') || 'dashboard';
    saveLastTab(activeId);

    if (updateHash) {
      history.replaceState(null, '', '#' + activeId);
    }

    if (focus) btn.focus();
  }

  function moveFocus(dir) {
    const current = document.activeElement;
    const idx = Math.max(0, tabs.indexOf(current));
    const next = (idx + dir + tabs.length) % tabs.length;
    tabs[next].focus();
  }

  // Click
  tabs.forEach(btn => {
    btn.addEventListener('click', () => {
      setActiveTab(btn.getAttribute('data-tab'), { updateHash: true, focus: false });
    });

    // Teclado (WAI-ARIA tabs)
    btn.addEventListener('keydown', (e) => {
      switch (e.key) {
        case 'ArrowLeft':
        case 'Left':
          e.preventDefault();
          moveFocus(-1);
          break;
        case 'ArrowRight':
        case 'Right':
          e.preventDefault();
          moveFocus(1);
          break;
        case 'Home':
          e.preventDefault();
          tabs[0].focus();
          break;
        case 'End':
          e.preventDefault();
          tabs[tabs.length - 1].focus();
          break;
        case 'Enter':
        case ' ':
          e.preventDefault();
          setActiveTab(btn.getAttribute('data-tab'), { updateHash: true, focus: true });
          break;
        default:
          break;
      }
    });
  });

  window.addEventListener('hashchange', () => {
    const h = getTabIdFromHash();
    if (h) setActiveTab(h, { updateHash: false, focus: false });
  });

  // Boot: hash > localStorage > default
  const initial = getTabIdFromHash() || loadLastTab() || tabs[0].getAttribute('data-tab');
  setActiveTab(initial, { updateHash: true, focus: false });
})();



