import { chromium, expect } from '../tools/node_modules/@playwright/test/index.mjs';
import { strict as assert } from 'node:assert';
import { existsSync } from 'node:fs';
import { mkdir } from 'node:fs/promises';
import { resolve, extname } from 'node:path';

assert.ok(existsSync('build/docs/index.html'), 'Documentation must build a static index page.');
const root = resolve('build/docs');
const server = Bun.serve({
  hostname: '127.0.0.1', port: 0,
  async fetch(request) {
    const pathname = decodeURIComponent(new URL(request.url).pathname);
    if (!pathname.startsWith('/minkexcel/')) return new Response('Not found', { status: 404 });
    const path = resolve(root, pathname.slice('/minkexcel/'.length) || 'index.html');
    if (!path.startsWith(root + '/')) return new Response('Not found', { status: 404 });
    const file = Bun.file(path);
    if (!await file.exists()) return new Response('Not found', { status: 404 });
    const types = { '.html': 'text/html', '.css': 'text/css', '.js': 'text/javascript', '.mjs': 'text/javascript', '.json': 'application/json' };
    return new Response(file, { headers: { 'Content-Type': types[extname(path)] || file.type } });
  },
});
const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  const errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  const base = `http://127.0.0.1:${server.port}/minkexcel/`;
  const llmResponse = await page.request.get(`${base}llms-full.txt`);
  assert.equal(llmResponse.status(), 200, 'LLM entry point is served under the GitHub Pages project path');
  assert.match(llmResponse.headers()['content-type'], /^text\/plain/);
  const llmText = await llmResponse.text();
  assert.match(llmText, /^# MinkExcel\n/);
  assert.match(llmText, /Self-contained documentation/);
  assert.match(llmText, /### Workbook\n/);
  assert.match(llmText, /### Worksheet\n/);
  assert.match(llmText, /### ReadLimits\n/);
  assert.match(llmText, /5,242,880/);
  assert.match(llmText, /readWorkbook\(bytes: Uint8Array/);
  assert.match(llmText, /export type ReadLimits = Partial/);
  assert.match(llmText, /Recorded results: Node/);
  assert.ok(!/<(?:svg|button|table|pre)\b/.test(llmText.replace(/```[\s\S]*?```/g, '')), 'Agent guide contains readable Markdown rather than interface markup');
  for (const filename of ['matrix.json', 'matrix-node.json']) {
    const dataset = await Bun.file(`benchmarks/${filename}`).json();
    const measurements = llmText.split(`## Recorded results: ${dataset.metadata.runtime}\n`)[1]?.split('\n## ')[0];
    assert.ok(measurements, 'Each recorded runtime has an embedded results table');
    for (const row of dataset.rows) {
      const expected = [row.workload, row.rows, ...['minkexcelExportMs', 'exceljsExportMs', 'minkexcelImportMs', 'exceljsImportMs'].map(key => row[key].toFixed(2)), row.minkexcelBytes, row.exceljsBytes].join(' | ');
      assert.ok(measurements.includes(`| ${expected} |`), `Embedded ${dataset.metadata.runtime} ${row.workload}/${row.rows} results match their recorded data`);
    }
  }
  const standardResponse = await page.request.get(`${base}llms.txt`);
  assert.equal(standardResponse.status(), 200);
  const llmsIndex = await standardResponse.text();
  assert.ok(llmsIndex.includes('](llms-full.txt)'), 'The index directs agents to the complete single-file documentation');
  assert.ok(llmsIndex.length < llmText.length, 'The index stays concise while the full guide contains all documentation');
  assert.equal((await page.request.get(`${base}llm.txt`)).status(), 404, 'Only the requested standard filenames are published');
  for (const match of (llmsIndex + '\n' + llmText).matchAll(/\[[^\]]+\]\(([^)]+)\)/g)) {
    const url = new URL(match[1], `${base}llms.txt`);
    if (url.origin === new URL(base).origin) assert.equal((await page.request.get(url.href)).status(), 200, `LLM reference resolves: ${match[1]}`);
  }
  for (const slug of ['index', 'getting-started', 'installation', 'api', 'comparison', 'benchmarks']) {
    await page.goto(`${base}${slug}.html`);
    await expect(page.locator('head link[rel="describedby"]')).toHaveAttribute('href', 'llms.txt');
    for (const example of await page.locator('pre code').allTextContents()) {
      assert.ok(llmText.includes(example), `Self-contained guide preserves the ${slug} example verbatim`);
    }
  }
  // Verify focus on rendered owners, including indicators inside clipping panels.
  for (const colorScheme of ['light', 'dark'] as const) {
    const context = await browser.newContext({ colorScheme, viewport: { width: 1440, height: 1000 } });
    const focusPage = await context.newPage();
    await focusPage.goto(`${base}benchmarks.html`);
    const rows = focusPage.getByRole('combobox', { name: 'Data rows', exact: true });
    await rows.waitFor();
    await focusPage.keyboard.press('Tab');
    await rows.focus();
    const indicator = await rows.evaluate(el => {
      const style = getComputedStyle(el);
      const color = style.outlineColor.match(/[\d.]+/g)!.slice(0, 3).map(Number);
      const background = getComputedStyle(document.body).backgroundColor.match(/[\d.]+/g)!.slice(0, 3).map(Number);
      const luminance = (rgb: number[]) => rgb.map(value => value / 255).map(value => value <= .04045 ? value / 12.92 : ((value + .055) / 1.055) ** 2.4).reduce((sum, value, i) => sum + value * [.2126, .7152, .0722][i], 0);
      const a = luminance(color), b = luminance(background);
      return { visible: el.matches(':focus-visible'), contrast: (Math.max(a, b) + .05) / (Math.min(a, b) + .05) };
    });
    assert.ok(indicator.visible && indicator.contrast >= 3, `${colorScheme}: dropdown keyboard ring contrasts with its surface`);
    for (const selector of ['#reproduce .copy', '#reproduce pre']) {
      const target = focusPage.locator(selector);
      await target.focus();
      const fits = await target.evaluate(el => {
        const style = getComputedStyle(el), box = el.getBoundingClientRect(), panel = el.closest('.code-block')!.getBoundingClientRect();
        const reach = parseFloat(style.outlineOffset) + parseFloat(style.outlineWidth);
        return el.matches(':focus-visible') && box.left - reach >= panel.left && box.right + reach <= panel.right && box.top - reach >= panel.top && box.bottom + reach <= panel.bottom;
      });
      assert.ok(fits, `${colorScheme}: ${selector} ring fits inside its clipping code panel`);
    }
    await focusPage.keyboard.press('Control+k');
    await focusPage.locator('.close-search').focus();
    assert.equal(await focusPage.locator('.search-input-row').evaluate(el => getComputedStyle(el).boxShadow), 'none', 'Search close has its own indicator without a duplicate field highlight');
    await focusPage.getByLabel('Search documentation', { exact: true }).fill('workbook');
    const lastResult = focusPage.locator('#search-results a').last();
    await lastResult.waitFor();
    await lastResult.focus();
    assert.ok(await lastResult.evaluate(el => {
      const style = getComputedStyle(el), box = el.getBoundingClientRect(), viewport = el.parentElement!.getBoundingClientRect();
      const reach = parseFloat(style.outlineOffset) + parseFloat(style.outlineWidth);
      return el.matches(':focus-visible') && box.top - reach >= viewport.top && box.bottom + reach <= viewport.bottom;
    }), 'Focused last search result keeps its ring inside the scrolling viewport');
    await context.close();
  }
  const forcedContext = await browser.newContext({ forcedColors: 'active', viewport: { width: 390, height: 844 } });
  const forcedPage = await forcedContext.newPage();
  await forcedPage.goto(`${base}index.html`);
  await forcedPage.keyboard.press('Control+k');
  assert.ok(await forcedPage.locator('.search-input-row').evaluate(el => {
    const style = getComputedStyle(el);
    return style.outlineStyle !== 'none' && parseFloat(style.outlineWidth) >= 2;
  }), 'Search field retains a visible outline when forced colors removes shadows');
  await forcedContext.close();
  let releaseIndex!: () => void;
  let pendingIndex: Promise<void> | undefined;
  const indexHeld = new Promise<void>(resolve => { releaseIndex = resolve; });
  await page.route('**/search.json', route => {
    pendingIndex = (async () => { await indexHeld; await route.continue(); })();
    return pendingIndex;
  });
  try {
    await page.goto(`${base}api.html`, { waitUntil: 'domcontentloaded' });
    await expect(page.getByRole('button', { name: 'Switch to dark theme' })).toBeVisible();
    await page.keyboard.press('Meta+k');
    await expect(page.locator('#search-dialog')).toBeVisible({ timeout: 1000 });
    await expect(page.getByLabel('Search documentation', { exact: true })).toBeFocused();
    await expect(page.locator('#search-results').getByText('Loading documentation search…', { exact: true })).toBeVisible();
    await page.getByLabel('Search documentation', { exact: true }).fill('ReadLimits');
    releaseIndex();
    await expect(page.locator('#search-results').getByRole('link', { name: 'ReadLimits · API reference', exact: true })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.locator('#search-dialog')).not.toBeVisible();
    await expect(page.getByRole('button', { name: 'Open documentation search' })).toBeFocused();
  } finally {
    releaseIndex();
    await pendingIndex;
    await page.unroute('**/search.json');
  }
  await page.route('**/search.json', route => route.fulfill({ status: 503, body: 'Unavailable' }));
  await page.goto(`${base}api.html`);
  await page.keyboard.press('Control+k');
  await expect(page.locator('#search-dialog')).toBeVisible();
  await expect(page.locator('#search-results').getByText('Search is unavailable. Use the documentation navigation.', { exact: true })).toBeVisible();
  await page.keyboard.press('Escape');
  await page.unroute('**/search.json');
  for (const [systemTheme, savedTheme, width] of [
    ['light', 'dark', 390], ['light', 'dark', 1440],
    ['dark', 'light', 390], ['dark', 'light', 1440],
  ] as const) {
    const context = await browser.newContext({ colorScheme: systemTheme, viewport: { width, height: width === 390 ? 844 : 1000 } });
    const switching = await context.newPage();
    switching.on('pageerror', error => errors.push(error.message));
    await switching.goto(`${base}index.html`);
    await switching.getByRole('button', { name: `Switch to ${savedTheme} theme` }).click();
    let releaseModule!: () => void;
    let pendingModule: Promise<void> | undefined;
    const moduleHeld = new Promise<void>(resolve => { releaseModule = resolve; });
    await switching.route('**/app.js', route => {
      pendingModule = (async () => { await moduleHeld; await route.continue(); })();
      return pendingModule;
    });
    try {
      await switching.goto(`${base}getting-started.html`, { waitUntil: 'commit' });
      await expect(switching.getByRole('heading', { name: 'Getting started', exact: true })).toBeVisible();
      assert.equal(await switching.locator('html').getAttribute('data-theme'), savedTheme, 'Saved theme applies before the application module loads');
      const expectedBackground = savedTheme === 'dark' ? 'rgb(10, 10, 10)' : 'rgb(255, 255, 255)';
      assert.equal(await switching.locator('body').evaluate(el => getComputedStyle(el).backgroundColor), expectedBackground, 'First painted page uses the saved theme');
      assert.equal(await switching.locator('html').evaluate(el => getComputedStyle(el).backgroundColor), expectedBackground, 'The document canvas matches the page theme');
      await expect(switching.getByRole('button', { name: 'Open documentation search' })).toBeVisible();
      const headerBefore = await switching.locator('.top-actions').boundingBox();
      const menuBefore = await switching.locator('.header-navigation').boundingBox();
      await mkdir('test-results', { recursive: true });
      await switching.evaluate(() => document.fonts.ready);
      await switching.screenshot({ path: `test-results/docs-navigation-${savedTheme}-${width}-before-module.png` });
      releaseModule();
      await expect(switching.locator('html')).toHaveClass(/navigation-ready/);
      await expect(switching.getByRole('button', { name: `Switch to ${systemTheme} theme` })).toBeVisible();
      assert.deepEqual(await switching.locator('.top-actions').boundingBox(), headerBefore, 'Navbar geometry stays fixed when enhancements initialize');
      assert.deepEqual(await switching.locator('.header-navigation').boundingBox(), menuBefore, 'Navbar menu keeps its position when enhanced navigation initializes');
      await switching.screenshot({ path: `test-results/docs-navigation-${savedTheme}-${width}-after-module.png` });
    } finally {
      releaseModule();
      await pendingModule;
      await context.close();
    }
  }
  await page.setViewportSize({ width: 320, height: 740 });
  await page.goto(`${base}getting-started.html`);
  assert.ok(await page.locator('h1').evaluate(el => parseFloat(getComputedStyle(el).fontSize) <= 34), 'Mobile documentation uses a compact page title');
  assert.ok(await page.locator('main p').first().evaluate(el => parseFloat(getComputedStyle(el).fontSize) >= 16), 'Mobile body text is readable');
  const menuButton = page.getByRole('button', { name: 'Open navigation' });
  assert.ok(await menuButton.evaluate(el => Boolean(el.closest('.topbar'))), 'Mobile navigation opens from the top navbar');
  assert.equal(await menuButton.evaluate(el => el.getBoundingClientRect().height), 32, 'Menu matches the user-requested compact navbar controls');
  assert.equal(await page.locator('.mobile-bar').count(), 0, 'Mobile navigation does not add a second header row');
  assert.ok(await page.locator('.topbar').evaluate(el => el.scrollWidth <= el.clientWidth), 'Navbar fits at 320px');
  await page.setViewportSize({ width: 1280, height: 900 });
  for (const name of ['index', 'getting-started', 'installation', 'api', 'comparison', 'benchmarks']) {
    await page.goto(`${base}${name}.html`);
    assert.equal(await page.locator('h1').count(), 1, `${name} has a main heading`);
    for (const href of await page.locator('a[href]').evaluateAll(links => links.map(link => link.getAttribute('href')))) {
      if (!href || /^(https?:|mailto:)/.test(href)) continue;
      const url = new URL(href, page.url());
      const response = await page.request.get(url.href);
      assert.equal(response.status(), 200, `Local link resolves: ${url.href}`);
      if (url.hash && url.pathname === new URL(page.url()).pathname) {
        assert.ok(await page.locator(`[id="${url.hash.slice(1)}"]`).count(), `Anchor exists: ${href}`);
      }
    }
  }
  await page.goto(`${base}benchmarks.html`);
  const chooseBenchmark = async (id: string, name: string, value: string) => {
    const text = await page.locator(`#${id} option[value="${value}"]`).textContent();
    await page.getByRole('combobox', { name, exact: true }).click();
    await page.getByRole('listbox', { name, exact: true }).getByRole('option', { name: text!, exact: true }).click();
  };
  const rowTrigger = page.getByRole('combobox', { name: 'Data rows', exact: true });
  await rowTrigger.click();
  const rowMenu = page.getByRole('listbox', { name: 'Data rows', exact: true });
  await expect(rowMenu).toBeVisible();
  await expect(rowMenu.getByRole('option', { name: '10,000', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(rowTrigger).toBeFocused();
  await page.keyboard.press('Home');
  await page.keyboard.press('Enter');
  await expect(page.locator('#chart-title')).toHaveText('Export · 100 data rows');
  await expect(rowMenu).not.toBeVisible();
  await expect(rowTrigger).toBeFocused();
  await rowTrigger.press('ArrowDown');
  await page.keyboard.press('End');
  await page.keyboard.press('Escape');
  await expect(page.locator('#row-count')).toHaveValue('100');
  await expect(rowTrigger).toHaveAttribute('aria-expanded', 'false');
  await chooseBenchmark('row-count', 'Data rows', '10000');
  const operationTrigger = page.getByRole('combobox', { name: 'Operation', exact: true });
  await operationTrigger.focus();
  await page.keyboard.press('i');
  await page.keyboard.press('Tab');
  await expect(page.locator('#chart-title')).toHaveText('Import · 10,000 data rows');
  await expect(rowTrigger).toBeFocused();
  await chooseBenchmark('operation', 'Operation', 'Export');
  await rowTrigger.click();
  await page.locator('#timings h2').click();
  await expect(rowMenu).not.toBeVisible();
  await expect(rowTrigger).toHaveAttribute('aria-expanded', 'false');
  await page.setViewportSize({ width: 390, height: 844 });
  await rowTrigger.evaluate(el => window.scrollTo({ top: el.getBoundingClientRect().top + scrollY - innerHeight + 60, behavior: 'instant' }));
  await rowTrigger.click();
  const menuBounds = await rowMenu.boundingBox();
  assert.ok(menuBounds && menuBounds.x >= 0 && menuBounds.y >= 0 && menuBounds.x + menuBounds.width <= 390 && menuBounds.y + menuBounds.height <= 844, 'Open mobile selector stays inside the viewport near its bottom edge');
  await rowMenu.getByRole('option', { name: '1,000', exact: true }).click();
  await expect(page.locator('#chart-title')).toHaveText('Export · 1,000 data rows');
  await chooseBenchmark('row-count', 'Data rows', '10000');
  await page.setViewportSize({ width: 1280, height: 900 });
  const fallbackContext = await browser.newContext();
  await fallbackContext.addInitScript(() => { Object.defineProperty(HTMLElement.prototype, 'showPopover', { value: undefined }); });
  const fallback = await fallbackContext.newPage();
  await fallback.goto(`${base}benchmarks.html`);
  await expect(fallback.getByLabel('Runtime')).toBeVisible();
  await fallback.getByLabel('Runtime').selectOption('node');
  await expect(fallback.locator('#chart-caption')).toContainText('Node');
  await fallbackContext.close();
  const bun = await Bun.file('benchmarks/matrix.json').json();
  const node = await Bun.file('benchmarks/matrix-node.json').json();
  const verify = async (source, operation) => {
    const rows = source.rows.filter(row => row.rows === 10000);
    const values = await page.locator('#timing-chart [data-value]').evaluateAll(elements => elements.map(element => Number(element.getAttribute('data-value'))));
    assert.deepEqual(values, rows.flatMap(row => [row[`minkexcel${operation}Ms`], row[`exceljs${operation}Ms`]]));
    assert.equal(await page.locator('#result-table tbody tr').count(), 9);
  };
  await verify(bun, 'Export');
  await chooseBenchmark('runtime', 'Runtime', 'node');
  await verify(node, 'Export');
  await chooseBenchmark('operation', 'Operation', 'Import');
  await verify(node, 'Import');
  for (const [runtime, source] of [['bun', bun], ['node', node]] as const) {
    await chooseBenchmark('runtime', 'Runtime', runtime);
    for (const operation of ['Export', 'Import']) {
      await chooseBenchmark('operation', 'Operation', operation);
      for (const count of [100, 1000, 10000]) {
        await chooseBenchmark('row-count', 'Data rows', String(count));
        const expected = source.rows.filter(row => row.rows === count).flatMap(row => [row[`minkexcel${operation}Ms`], row[`exceljs${operation}Ms`]]);
        const actual = await page.locator('#timing-chart [data-value]').evaluateAll(elements => elements.map(element => Number(element.getAttribute('data-value'))));
        assert.deepEqual(actual, expected);
        assert.ok(Number(await page.locator('#timing-chart [data-percent]').last().getAttribute('data-percent')) > 15, 'Chart uses a readable scale at every size');
      }
    }
  }
  await page.getByRole('button', { name: 'Open documentation search' }).click();
  await page.getByLabel('Search documentation').fill('ReadLimits');
  await page.locator('#search-results a').first().click();
  assert.ok(page.url().endsWith('api.html#read-limits'));
  await page.goto(`${base}index.html`);
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.getByRole('button', { name: 'Copy code' }).first().click();
  assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'npm install minkexcel');
  await page.setViewportSize({ width: 390, height: 844 });
  await page.getByRole('button', { name: 'Open navigation' }).click();
  assert.ok(await page.locator('#mobile-navigation').evaluate(el => el.open), 'Mobile navigation opens as a dialog');
  await page.keyboard.press('Escape');
  await expect(page.getByRole('button', { name: 'Open navigation' })).toHaveAttribute('aria-expanded', 'false');
  await page.getByRole('button', { name: 'Open navigation' }).click();
  await page.locator('#mobile-navigation').getByRole('link', { name: 'API reference', exact: true }).click();
  assert.ok(page.url().endsWith('api.html'));
  assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), 'Mobile page does not overflow');
  await mkdir('test-results', { recursive: true });
  await page.screenshot({ path: 'test-results/docs-api-mobile.png', fullPage: false });
  for (const width of [320, 390, 768]) {
    await page.setViewportSize({ width, height: 844 });
    for (const name of ['index', 'getting-started', 'installation', 'api', 'comparison', 'benchmarks']) {
      await page.goto(`${base}${name}.html`);
      assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), `${name} fits at ${width}px`);
      if (name === 'api') assert.ok(await page.locator('#workbook td').first().isVisible(), 'API behavior is readable on narrow screens');
    }
  }
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto(`${base}getting-started.html`);
  await page.screenshot({ path: 'test-results/docs-mobile.png', fullPage: false });
  await page.locator('.mobile-toc summary').click();
  await page.locator('.mobile-toc').getByRole('link', { name: 'Import and read values', exact: true }).click();
  assert.ok(page.url().endsWith('#read'));
  await page.keyboard.press('Control+k');
  assert.ok(await page.locator('#search-dialog').evaluate(el => el.open));
  await page.keyboard.press('Escape');
  await expect(page.locator('#search-dialog')).not.toBeVisible();
  await page.getByRole('button', { name: 'Switch to dark theme' }).click();
  await page.reload();
  assert.equal(await page.locator('html').getAttribute('data-theme'), 'dark');
  await page.getByRole('button', { name: 'Switch to light theme' }).click();
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.goto(`${base}benchmarks.html`);
  await page.screenshot({ path: 'test-results/docs-benchmarks.png', fullPage: true });
  await page.goto(`${base}index.html`);
  await page.screenshot({ path: 'test-results/docs-home.png', fullPage: false });
  await page.goto(`${base}getting-started.html`);
  await page.screenshot({ path: 'test-results/docs-desktop.png', fullPage: false });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto(`${base}benchmarks.html`);
  await page.locator('#timings').screenshot({ path: 'test-results/docs-chart-mobile.png' });
  assert.deepEqual(errors, []);
  const plain = await browser.newPage({ javaScriptEnabled: false });
  await plain.goto(`${base}benchmarks.html`);
  assert.equal(await plain.locator('#timing-chart [data-value]').count(), 6, 'Chart is present without JavaScript');
  assert.equal(await plain.locator('#result-table tbody tr').count(), 9, 'Results are readable without JavaScript');
  await plain.setViewportSize({ width: 390, height: 844 });
  await plain.locator('.fallback-navigation summary').click();
  assert.ok(await plain.locator('.fallback-navigation').getByRole('link', { name: 'API reference', exact: true }).isVisible(), 'Navigation is accessible on mobile without JavaScript');
  console.log('Documentation: all pages, local links, charts, search, copy, mobile, pre-paint navigation and no-JS checks passed.');
} finally {
  await browser.close();
  server.stop(true);
}
