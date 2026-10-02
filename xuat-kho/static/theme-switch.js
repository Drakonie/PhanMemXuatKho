/* Change presentation in place, keeping uploads, edits and review state intact. */
(() => {
  'use strict';
  const key = 'xuat-kho-theme';
  const root = document.documentElement;
  let preference = 'linh';
  try {
    const saved = window.localStorage.getItem(key);
    if (saved === 'normal' || saved === 'linh') preference = saved;
  } catch (_) { /* Private browsing can make storage unavailable. */ }

  function apply(theme, persist = false) {
    preference = theme === 'normal' ? 'normal' : 'linh';
    root.dataset.theme = preference;
    const stylesheet = document.getElementById('linh-theme');
    if (stylesheet) stylesheet.disabled = preference !== 'linh';
    document.title = preference === 'linh'
      ? 'Linh linh yêu · Xuất kho Huyền An 68'
      : 'Xuất kho theo template · Huyền An 68';
    const color = document.querySelector('meta[name="theme-color"]');
    if (color) color.content = preference === 'linh' ? '#fff3f8' : '#ffffff';
    const brand = document.querySelector('.brand');
    if (brand) brand.setAttribute('aria-label', preference === 'linh'
      ? 'Linh linh yêu · Xuất kho Huyền An 68' : 'Xuất kho Huyền An 68');
    ['normal', 'linh'].forEach(name => {
      const button = document.getElementById(`theme-${name}`);
      if (button) button.setAttribute('aria-pressed', String(preference === name));
    });
    if (persist) {
      try { window.localStorage.setItem(key, preference); } catch (_) { /* The toggle still works. */ }
    }
  }

  // This script runs in the head so a saved normal theme takes effect before paint.
  apply(preference);
  function ready() {
    apply(preference);
    ['normal', 'linh'].forEach(name => {
      document.getElementById(`theme-${name}`).addEventListener('click', () => apply(name, true));
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', ready, {once: true});
  else ready();
})();
