const select = selector => document.querySelector(selector);
document.documentElement.classList.add('js');
const menu = select('.menu-toggle');
const navigation = select('#mobile-navigation');
const searchDialog = select('#search-dialog');
const input = select('#search');
const results = select('#search-results');
menu.hidden = false;
menu.addEventListener('click', () => {
  navigation.showModal();
  menu.setAttribute('aria-expanded', 'true');
});
select('.close-navigation').addEventListener('click', () => navigation.close());
navigation.addEventListener('close', () => {
  menu.setAttribute('aria-expanded', 'false');
  menu.focus();
});
navigation.addEventListener('click', event => {
  if (event.target === navigation) {
    const bounds = navigation.getBoundingClientRect();
    if (event.clientX > bounds.right || event.clientX < bounds.left) navigation.close();
  }
  if (event.target.closest('a')) navigation.close();
});
matchMedia('(min-width: 900px)').addEventListener('change', event => {
  if (event.matches && navigation.open) navigation.close();
});
document.documentElement.classList.add('navigation-ready');
const theme = select('.theme-toggle');
theme.hidden = false;
function dark() {
  return document.documentElement.dataset.theme
    ? document.documentElement.dataset.theme === 'dark'
    : matchMedia('(prefers-color-scheme: dark)').matches;
}
function applyTheme(value) {
  if (value) document.documentElement.dataset.theme = value;
  theme.setAttribute('aria-label', `Switch to ${dark() ? 'light' : 'dark'} theme`);
}
try { applyTheme(localStorage.getItem('minkexcel-theme')); } catch { applyTheme(); }
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => applyTheme());
theme.addEventListener('click', () => {
  const value = dark() ? 'light' : 'dark';
  applyTheme(value);
  try { localStorage.setItem('minkexcel-theme', value); } catch { /* Storage may be unavailable. */ }
});
let toastTimer;
document.querySelectorAll('.copy').forEach(button => {
  button.hidden = false;
  button.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(button.closest('.code-block').querySelector('code').textContent);
      select('#copy-status').textContent = 'Code copied to clipboard';
    } catch {
      select('#copy-status').textContent = 'Select the code and copy it manually.';
    }
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { select('#copy-status').textContent = ''; }, 2400);
  });
});
let searchData;
let searchFailed = false;
const searchTrigger = select('.search-trigger');
searchTrigger.hidden = false;
function openSearch() {
  if (navigation.open) navigation.close();
  if (!searchDialog.open) searchDialog.showModal();
  input.focus();
}
searchTrigger.addEventListener('click', openSearch);
select('.close-search').addEventListener('click', () => searchDialog.close());
searchDialog.addEventListener('click', event => {
  if (event.target === searchDialog) {
    const bounds = searchDialog.getBoundingClientRect();
    if (event.clientX < bounds.left || event.clientX > bounds.right || event.clientY < bounds.top || event.clientY > bounds.bottom) searchDialog.close();
  }
});
searchDialog.addEventListener('close', () => searchTrigger?.focus());
function renderSearch() {
  const query = input.value.trim().toLowerCase();
  results.replaceChildren();
  results.setAttribute('aria-busy', String(!searchData && !searchFailed));
  if (!searchData) {
    const message = document.createElement('p');
    message.textContent = searchFailed
      ? 'Search is unavailable. Use the documentation navigation.'
      : 'Loading documentation search…';
    results.append(message);
    select('#search-status').textContent = message.textContent;
    return;
  }
  if (!query) {
    results.innerHTML = '<p>Search guides, API methods and benchmarks.</p>';
    select('#search-status').textContent = '';
    return;
  }
  const matches = searchData.filter(item => item.title.toLowerCase().includes(query) || item.text.toLowerCase().includes(query))
    .sort((a, b) => Number(b.title.toLowerCase().includes(query)) - Number(a.title.toLowerCase().includes(query))).slice(0, 8);
  for (const item of matches) {
    const link = document.createElement('a');
    link.href = new URL(item.href, import.meta.url).href;
    link.textContent = item.title;
    link.addEventListener('click', () => searchDialog.close());
    results.append(link);
  }
  if (!matches.length) {
    const empty = document.createElement('p');
    empty.textContent = 'No results. Try “workbook”, “limits” or “export”.';
    results.append(empty);
  }
  select('#search-status').textContent = `${matches.length} search results`;
}
input.addEventListener('input', renderSearch);
input.addEventListener('keydown', event => {
  if (event.key === 'ArrowDown') { event.preventDefault(); results.querySelector('a')?.focus(); }
  if (event.key === 'Enter') results.querySelector('a')?.click();
});
results.addEventListener('keydown', event => {
  const links = [...results.querySelectorAll('a')];
  const index = links.indexOf(document.activeElement);
  if (event.key === 'ArrowDown') { event.preventDefault(); links[Math.min(index + 1, links.length - 1)]?.focus(); }
  if (event.key === 'ArrowUp') { event.preventDefault(); index > 0 ? links[index - 1].focus() : input.focus(); }
});
window.addEventListener('keydown', event => {
  if (event.isComposing) return;
  if (event.key === 'Escape' && searchDialog.open) {
    event.preventDefault();
    searchDialog.close();
    return;
  }
  const active = document.activeElement;
  const editing = /INPUT|TEXTAREA|SELECT/.test(active.tagName) || active.isContentEditable;
  const command = (event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey;
  const searchShortcut = command && (event.code === 'KeyK' || event.key.toLowerCase() === 'k');
  const slashShortcut = event.key === '/' && !editing && !event.ctrlKey && !event.metaKey && !event.altKey;
  if (searchShortcut || slashShortcut) {
    event.preventDefault();
    openSearch();
  }
}, { capture: true });
renderSearch();
async function loadSearch() {
  try {
    const response = await fetch(new URL('search.json', import.meta.url));
    if (!response.ok) throw new Error('Search index unavailable');
    searchData = await response.json();
  } catch { searchFailed = true; }
  renderSearch();
}
loadSearch();
document.querySelectorAll('.mobile-toc a').forEach(link => {
  link.addEventListener('click', () => { select('.mobile-toc').open = false; });
});
const controls = select('.chart-controls');
if (controls) {
  try {
    const [{ timingChart, sizeChart, resultsTable }, datasets, { enhanceSelects }, { numericText }] = await Promise.all([
      import('./charts.mjs'),
      Promise.all(['matrix.json', 'matrix-node.json'].map(async name => {
        const response = await fetch(new URL(`data/${name}`, import.meta.url));
        if (!response.ok) throw new Error('Benchmark data unavailable');
        return response.json();
      })),
      import('./select.js'),
      import('./numbers.mjs'),
    ]);
    const update = () => {
      const data = datasets[select('#runtime').value === 'bun' ? 0 : 1];
      const operation = select('#operation').value;
      const count = Number(select('#row-count').value);
      const rows = data.rows.filter(row => row.rows === count);
      select('#timing-chart').innerHTML = timingChart(rows, operation);
      select('#size-chart').innerHTML = sizeChart(rows);
      select('#chart-title').innerHTML = numericText(`${operation} · ${count.toLocaleString('en-US')} data rows`);
      select('#chart-caption').innerHTML = numericText(`${data.metadata.runtime} · medians of ${data.metadata.iterations} samples after ${data.metadata.warmups} warmup · ${data.metadata.columns} columns plus a header`);
      select('#result-table').innerHTML = resultsTable(data.rows);
      select('#table-caption').innerHTML = numericText(`${data.metadata.runtime} · all workloads and row counts. Times in milliseconds.`);
    };
    controls.addEventListener('change', update);
    enhanceSelects(controls);
    controls.hidden = false;
  } catch { select('#chart-caption').append(' · Interactive data could not load; default recorded results are shown.'); }
}
if ('IntersectionObserver' in window) {
  const observer = new IntersectionObserver(entries => {
    for (const entry of entries) if (entry.isIntersecting) {
      document.querySelectorAll('.toc a, .api-nav a').forEach(link => {
        const active = link.hash === `#${entry.target.id}` && new URL(link.href).pathname === location.pathname;
        link.classList.toggle('active', active);
        if (active) link.setAttribute('aria-current', 'location');
        else link.removeAttribute('aria-current');
      });
    }
  }, { rootMargin: '-15% 0px -65% 0px' });
  document.querySelectorAll('main section[id]').forEach(section => observer.observe(section));
}
