// Runs synchronously in the document head, before styles or application modules.
(() => {
  const root = document.documentElement;
  root.classList.add('js');
  root.dataset.platform = /mac|iphone|ipad/i.test(navigator.userAgentData?.platform || navigator.platform)
    ? 'apple' : 'other';
  try {
    const saved = localStorage.getItem('minkexcel-theme');
    if (saved === 'light' || saved === 'dark') root.dataset.theme = saved;
  } catch { /* CSS follows the system theme when storage is unavailable. */ }
})();
