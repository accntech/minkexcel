# MinkExcel documentation site

The site covers project motivation, getting started, installation, the public
API, ExcelJS migration/comparison and recorded benchmarks. It uses static HTML,
CSS and small ES modules with no site build dependencies. Existing contributor
tools supply the optional browser verification harness.

## Build and preview

From the repository root, with Node.js installed:

```sh
npm run docs:build
npm run docs:preview
```

Open http://127.0.0.1:4173/minkexcel/. Set `DOCS_PORT` to use another port.
Build output goes to `build/docs/`, which is ignored by Git. The preview serves
the site at the same `/minkexcel/` path as a GitHub project site. Rebuild and
refresh after editing. Documentation pages use directory URLs such as
`/minkexcel/getting-started/`, backed by `getting-started/index.html`. The home
page uses `/minkexcel/`; explicit `index.html` URLs redirect to the containing
directory when JavaScript is enabled. Navigation,
search and agent references use these clean URLs. Previous `.html` page URLs
redirect to their directory equivalents; JavaScript preserves query strings and
section anchors, with a meta-refresh and link fallback when JavaScript is disabled.
The preview also redirects directory requests without a trailing slash.
Serve over HTTP; interactive ES modules need an HTTP
origin rather than a `file://` URL.

All page content and initial bar graphs are built into HTML. JavaScript adds
search, code copying, a persisted theme, mobile navigation and benchmark
controls. Content, navigation, default charts and result tables remain available
without JavaScript. The site self-hosts Geist Sans and Geist Mono and makes no
external asset requests. Solar Linear utility icons and the official GitHub brand mark are embedded as SVGs.

The layout starts with a single reading column at 320px. API rows stack on
phones, and comparison tables scroll within their own region. A persistent
sidebar appears at 900px, with a right section list at 1280px. Graph labels and
values remain readable on phones. Below 900px, a compact menu button lives in
the top navbar; its drawer and no-JavaScript disclosure share the same slot,
without an additional menu row.

## Publish on GitHub Pages

1. In `accntech/minkexcel`, open **Settings → Pages**.
2. Set **Build and deployment → Source** to **GitHub Actions**.
3. Push the documentation changes to `main`, or run the **Documentation**
   workflow manually on `main` after the changes have been pushed.
4. The workflow builds and checks the site, uploads `build/docs/`, and publishes
   it to https://accntech.github.io/minkexcel/.

Pull requests build and run browser checks without deploying. Deployment uses
the `github-pages` environment and requires Pages and ID token write permission;
the build job has read-only repository access. If environment protection requires
approval, approve the deployment in GitHub. The configured workflow follows
[GitHub's custom workflow documentation](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages).

Generated HTML and assets use relative links, so the output can also be served
at another subdirectory or on another static host without modifying a base URL.

The build publishes `/minkexcel/llms.txt` as a concise index that links to
`/minkexcel/llms-full.txt`, the complete self-contained Markdown guide. For the
full guide, the build combines the introduction in `llms.txt` with all six pages
from `content.mjs`, preserving complete code examples,
API tables and public types. It also includes public exports, the expanded
`ReadLimits` type, and every recorded benchmark median and compressed file size
for Bun and Node.js. Version, exports and measurements come from the package and
source datasets, so rebuilding keeps the guide synchronized. Links are optional
references; agents can use the guide without fetching other documentation.
Both text files are linked as “Documentation index” and “Complete documentation”
in the sidebar's “For AI agents” group and mobile
navigation, including without JavaScript. Every HTML page also links to `llms.txt` with `rel="describedby"`, following
the [llms.txt proposal](https://llmstxt.org/). Local references stay relative to the
published documentation path. Review this summary when public behavior changes.

## Edit and refresh

- `content.mjs`: the six documentation pages and examples.
- `llms.txt`: maintained agent index, key conventions and optional references; links to the full guide.
- `llms.mjs`: generates `llms-full.txt` from the shared documentation, public declarations and recorded benchmark tables.
- `styles.css`: layout, responsive rules, light/dark palettes and print styles.
- `app.js`: optional progressive enhancements.
- `select.js`: benchmark select-only comboboxes, styled option popovers, keyboard handling and native-select fallback; loaded with chart enhancements only on the benchmarks page.
- `boot.js`: inlined before styles to restore the saved theme and platform keycap before the first paint.
- `charts.mjs`: responsive HTML bar graph/table renderers used during build and in-browser.
- `numbers.mjs`: shared numeric text/markup helpers for static content, chart updates and enhanced option labels; wraps numeric literals in `.numeric` for Geist Mono and tabular numerals while preserving code, identifiers, SVGs, HTML entities and native options. Numeric-only native data-row options use the selector's mono style.
- `build.mjs`: page shell, navigation, search index and asset assembly.
- `icons.mjs`: embedded Solar Linear utility SVGs, the official GitHub brand mark, and the shared icon helper.
- `assets/`: self-hosted fonts and asset licenses, plus Solar attribution.
- `serve.mjs`: local static preview server.

Public formatting/value type definitions come directly from `src/model.ts`.
The version comes from `package.json`. Benchmarks come from
`benchmarks/matrix.json` and `benchmarks/matrix-node.json`, including raw samples;
the builder copies the recorded files unchanged into the site. No performance
data is fabricated and building never reruns timing benchmarks.

When updating measurements, review runtime/version/date descriptions and bundle
size claims in `content.mjs` as well as the new JSON. API member descriptions are
maintained in `content.mjs`; review them when library behavior changes. The
benchmarks page also embeds the template workflow results from
`benchmarks/template.json` and `benchmarks/template-node.json` and publishes both
datasets for download. Template mode and ReadOptions are included in the API
reference and the generated agent guide.

## Verify

Install tools and Chromium as described in the contributor guide, then run:

```sh
npm run test:docs
```

The harness verifies all six pages and local links under the GitHub project path,
clean directory URLs, legacy `.html` redirects, both LLM entry-point filenames,
their local references and HTML discovery links,
chart values against the recorded datasets, runtime/operation/row switching,
search navigation, clipboard copying, mobile layout and navigation, no-JavaScript
content, and browser errors. Screenshots go to `test-results/docs-*.png`.
Navigation checks delay the application module to verify the saved theme and
header geometry before and after initialization, on phones and desktops.
Benchmark selector checks cover keyboard commit/cancel and typeahead, pointer
selection and outside dismissal, accessible state, mobile popup positioning
and the native-select fallback.

## Design assets

Geist Sans and Geist Mono come from the official `geist` package (1.7.2) and use
the SIL Open Font License in `assets/licenses/Geist-OFL.txt`.
Solar Linear SVGs come from `@iconify-json/solar` (1.2.13). Solar is by
[480 Design](https://www.figma.com/community/file/1166831539721848736) under
CC BY 4.0; attribution is visible in the site footer and recorded in
`assets/licenses/Solar-ATTRIBUTION.md`.

The GitHub repository link uses the official [Primer Octicons `mark-github-24`](https://github.com/primer/octicons/blob/main/icons/mark-github-24.svg) SVG, embedded in `icons.mjs` with `currentColor`. Its MIT license is recorded in `assets/licenses/Octicons-LICENSE.txt`.
