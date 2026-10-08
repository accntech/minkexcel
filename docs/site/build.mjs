import { readFile, writeFile, mkdir, copyFile, rm, cp } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { pages, escape } from './content.mjs';
import { icon, githubMark } from './icons.mjs';
import { timingChart, sizeChart, resultsTable } from './charts.mjs';
import { numericMarkup, numericText } from './numbers.mjs';
import { llmsFullGuide } from './llms.mjs';

const root = fileURLToPath(new URL('../../', import.meta.url));
const output = join(root, 'build/docs');
const [manifest, bun, node, model, boot, templateBun, templateNode] = await Promise.all([
  readFile(join(root, 'package.json'), 'utf8').then(JSON.parse),
  readFile(join(root, 'benchmarks/matrix.json'), 'utf8').then(JSON.parse),
  readFile(join(root, 'benchmarks/matrix-node.json'), 'utf8').then(JSON.parse),
  readFile(join(root, 'src/model.ts'), 'utf8'),
  readFile(join(root, 'docs/site/boot.js'), 'utf8'),
  readFile(join(root, 'benchmarks/template.json'), 'utf8').then(JSON.parse),
  readFile(join(root, 'benchmarks/template-node.json'), 'utf8').then(JSON.parse),
]);
const documents = pages({
  version: manifest.version, bun, node, templateBun, templateNode, timingChart, sizeChart, resultsTable,
  types: model.slice(0, model.indexOf('type Address')).replaceAll('export ', '').trim(),
});
const [introduction, publicExports, reader] = await Promise.all([
  readFile(join(root, 'docs/site/llms.txt'), 'utf8'),
  readFile(join(root, 'src/index.ts'), 'utf8'),
  readFile(join(root, 'src/read.ts'), 'utf8'),
]);
const llmsFull = llmsFullGuide({
  introduction, version: manifest.version, exports: publicExports,
  limitFields: [...reader.match(/const defaultLimits: Limits = \{([\s\S]*?)\n\};/)[1].matchAll(/(\w+):/g)].map(match => match[1]),
  datasets: [bun, node],
  documents: pages({ version: manifest.version, bun, node, templateBun, templateNode, types: model.slice(0, model.indexOf('type Address')).replaceAll('export ', '').trim(), timingChart: () => '', sizeChart: () => '', resultsTable: () => '' }),
});
await rm(output, { recursive: true, force: true });
await mkdir(join(output, 'data'), { recursive: true });
await Promise.all([
  ...['styles.css', 'app.js', 'charts.mjs', 'select.js', 'numbers.mjs'].map(name => copyFile(join(root, 'docs/site', name), join(output, name))),
  writeFile(join(output, 'llms.txt'), introduction),
  writeFile(join(output, 'llms-full.txt'), llmsFull),
  cp(join(root, 'docs/site/assets'), join(output, 'assets'), { recursive: true }),
  copyFile(join(root, 'assets/icon.png'), join(output, 'icon.png')),
  copyFile(join(root, 'benchmarks/matrix.json'), join(output, 'data/matrix.json')),
  copyFile(join(root, 'benchmarks/matrix-node.json'), join(output, 'data/matrix-node.json')),
  copyFile(join(root, 'benchmarks/template.json'), join(output, 'data/template.json')),
  copyFile(join(root, 'benchmarks/template-node.json'), join(output, 'data/template-node.json')),
  writeFile(join(output, '.nojekyll'), ''),
]);
const apiLinks = [
  ['io', 'Read and write'], ['templates', 'Edit a template'], ['workbook', 'Workbook'], ['worksheet', 'Worksheet'],
  ['row', 'Row'], ['cell', 'Cell'], ['column', 'Column'], ['types', 'Types'],
  ['read-limits', 'Import limits'], ['errors', 'Errors and cancellation'],
];
const pageIcons = {
  index: 'book-bookmark', 'getting-started': 'document-text', installation: 'download',
  api: 'code-square', comparison: 'transfer-horizontal', benchmarks: 'chart-2',
};
const pageHref = slug => slug === 'index' ? './' : `${slug}/`;
const search = [];
for (const [index, page] of documents.entries()) {
  const headings = [...page.body.matchAll(/<section id="([^"]+)"><h2>([^<]+)/g)]
    .map(match => ({ id: match[1], title: match[2] }));
  const nav = ['Overview', 'Guides', 'Reference', 'Measurements'].map(group => `
    <div class="nav-group"><p>${group}</p>
      ${documents.filter(doc => doc.category === group).map(doc => `
        <a href="${pageHref(doc.slug)}" ${doc.slug === page.slug ? 'aria-current="page"' : ''}>
          ${icon(pageIcons[doc.slug])}<span>${doc.label}</span>
        </a>
        ${doc.slug === 'api' ? `<div class="api-nav">${apiLinks.map(([id, title]) => `
          <a href="api/#${id}" data-section="${id}">${title}</a>`).join('')}
        </div>` : ''}`).join('')}
    </div>`).join('') + `
    <div class="nav-group"><p>For AI agents</p>
      <a href="llms.txt" type="text/plain" title="llms.txt">${icon('document-text')}<span>Documentation index</span></a>
      <a href="llms-full.txt" type="text/plain" title="llms-full.txt">${icon('document-text')}<span>Complete documentation</span></a>
    </div>`;
  const sectionLinks = headings.map(h => `<a href="#${h.id}">${h.title}</a>`).join('');
  const plain = page.body.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
  search.push({ title: page.label, href: pageHref(page.slug), text: plain });
  for (const heading of headings) {
    search.push({ title: `${heading.title} · ${page.label}`, href: `${pageHref(page.slug)}#${heading.id}`, text: heading.title });
  }
  const previous = documents[index - 1], next = documents[index + 1];
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light dark"><meta name="description" content="${escape(page.description)}">
  <meta property="og:title" content="${escape(page.label)} — MinkExcel">
  <meta property="og:description" content="${escape(page.description)}"><meta property="og:type" content="website">
  <title>${escape(page.label)} — MinkExcel documentation</title>
  <link rel="describedby" href="llms.txt" type="text/plain">
  <script>if (location.pathname.endsWith('/index.html')) location.replace('./' + location.search + location.hash);
${boot}</script>
  <link rel="icon" href="icon.png">
  <link rel="preload" href="assets/fonts/Geist-Variable.woff2" as="font" type="font/woff2" crossorigin>
  <link rel="preload" href="assets/fonts/GeistMono-Variable.woff2" as="font" type="font/woff2" crossorigin>
  <link rel="stylesheet" href="styles.css"><script type="module" src="app.js"></script>
</head>
<body class="page-${page.slug}">
  <a class="skip-link" href="#main">Skip to content</a>
  <header class="topbar">
    <a class="brand" href="./"><img src="icon.png" width="36" height="36" alt=""><span class="brand-copy"><span class="brand-name"><strong>Mink</strong>Excel</span><span class="brand-subtitle">Documentation</span></span></a>
    <div class="top-actions">
      <button class="search-trigger" type="button" aria-label="Open documentation search" aria-haspopup="dialog" aria-controls="search-dialog" aria-keyshortcuts="Meta+K Control+K" hidden>
        ${icon('magnifier')}<span>Search documentation…</span><kbd aria-hidden="true"><span class="shortcut-modifier"><span class="shortcut-mac">⌘</span><span class="shortcut-control">Ctrl</span></span><span>K</span></kbd>
      </button>
      <a href="https://github.com/accntech/minkexcel" class="github-link" aria-label="GitHub repository" title="GitHub repository">${githubMark}</a>
      <button class="icon-button theme-toggle" type="button" aria-label="Switch to dark theme" hidden>
        ${icon('moon', 'moon-icon')}${icon('sun', 'sun-icon')}
      </button>
      <div class="header-navigation">
        <button class="icon-button menu-toggle" aria-label="Open navigation" aria-controls="mobile-navigation" aria-expanded="false" type="button" hidden>${icon('hamburger-menu')}</button>
        <details class="fallback-navigation"><summary aria-label="Documentation navigation">${icon('hamburger-menu')}</summary><nav aria-label="Documentation">${nav}</nav></details>
      </div>
    </div>
  </header>
  <div class="layout">
    <aside class="sidebar" id="navigation">
      <div class="sidebar-heading"><span>Documentation</span><span class="sidebar-version">v${escape(manifest.version)}</span></div>
      <nav aria-label="Documentation">${nav}</nav>
    </aside>
    <main id="main">
      <nav class="breadcrumb" aria-label="Breadcrumb"><a href="./">Documentation</a>${icon('alt-arrow-right')}<span>${page.label}</span></nav>
      <details class="mobile-toc"><summary>On this page ${icon('alt-arrow-down')}</summary><nav aria-label="Page sections">${sectionLinks}</nav></details>
      <div class="page-intro"><h1>${escape(page.title)}</h1><p class="lead">${numericText(page.description)}</p></div>
      ${numericMarkup(page.body)}
      <nav class="page-pagination" aria-label="Adjacent pages">
        ${previous ? `<a href="${pageHref(previous.slug)}"><span>${icon('arrow-left')} Previous</span><strong>${previous.label}</strong></a>` : '<div></div>'}
        ${next ? `<a href="${pageHref(next.slug)}"><span>Next ${icon('arrow-right')}</span><strong>${next.label}</strong></a>` : '<div></div>'}
      </nav>
      <footer class="footer">
        <div><span>MinkExcel · MIT License</span><a href="https://www.figma.com/community/file/1166831539721848736">Solar icons by 480 Design</a></div>
        <a href="https://github.com/accntech/minkexcel/blob/main/docs/site/content.mjs">Edit this page ${icon('arrow-right-up')}</a>
      </footer>
    </main>
    <aside class="toc"><p>On this page</p><nav aria-label="On this page">${sectionLinks}</nav>
      <a class="edit-link" href="https://github.com/accntech/minkexcel/blob/main/docs/site/content.mjs">${icon('document-text')} Edit this page</a>
    </aside>
  </div>
  <dialog id="mobile-navigation" class="nav-dialog" aria-labelledby="navigation-title">
    <div class="dialog-heading"><strong id="navigation-title">Documentation</strong><button type="button" class="icon-button close-navigation" aria-label="Close navigation">${icon('close-square')}</button></div>
    <nav aria-label="Documentation">${nav}</nav>
    <a class="drawer-github" href="https://github.com/accntech/minkexcel">View on GitHub ${icon('arrow-right-up')}</a>
  </dialog>
  <dialog id="search-dialog" class="search-dialog" aria-label="Documentation search">
    <div class="search-input-row">${icon('magnifier')}<label class="sr-only" for="search">Search documentation</label>
      <input type="search" id="search" placeholder="Search documentation…" autocomplete="off" aria-controls="search-results">
      <button type="button" class="icon-button close-search" aria-label="Close search">${icon('close-square')}</button>
    </div>
    <div id="search-results"><p>Search guides, API methods and benchmarks.</p></div>
    <div class="search-help"><span>Type to search</span><kbd>esc to close</kbd></div><span id="search-status" class="sr-only" role="status"></span>
  </dialog>
  <div id="copy-status" class="toast" role="status"></div>
</body>
</html>`;
  const directory = page.slug === 'index' ? output : join(output, page.slug);
  await mkdir(directory, { recursive: true });
  // Keep local assets and links relative to the site root from nested pages.
  const rendered = page.slug === 'index' ? html : html.replace(/\b(href|src)="(?![a-z]+:|\/|#)([^"]+)"/gi, (_, attribute, url) => `${attribute}="../${url}"`);
  await writeFile(join(directory, 'index.html'), rendered);
  if (page.slug !== 'index') {
    const target = pageHref(page.slug);
    await writeFile(join(output, `${page.slug}.html`), `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>${escape(page.label)} — MinkExcel</title>
<script>location.replace(${JSON.stringify(target)} + location.search + location.hash);</script>
<meta http-equiv="refresh" content="0;url=${target}"></head>
<body><a href="${target}">Continue to ${escape(page.label)}</a></body></html>`);
  }
}
await writeFile(join(output, 'search.json'), JSON.stringify(search));
console.log(`Built ${documents.length} documentation pages → build/docs`);
