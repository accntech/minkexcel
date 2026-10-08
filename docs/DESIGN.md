---
name: "MinkExcel documentation"
description: "Restrained read-mode documentation in the approved Geist direction."
colors:
  bg: "#ffffff"
  surface: "#fafafa"
  subtle: "#f3f3f3"
  ink: "#171717"
  muted: "#666666"
  line: "#e8e8e8"
  green: "#217346"
  accent: "#edf6f0"
  excel: "#7b8da5"
  string: "#216e46"
  keyword: "#7935a3"
  number: "#975116"
  focus: "#217346"
  dark-bg: "#0a0a0a"
  dark-surface: "#111111"
  dark-subtle: "#1a1a1a"
  dark-ink: "#ededed"
  dark-muted: "#a1a1a1"
  dark-line: "#282828"
  dark-green: "#75c899"
  dark-accent: "#14291c"
  dark-excel: "#91a6c4"
  dark-string: "#8ad8a6"
  dark-keyword: "#cf9fff"
  dark-number: "#f0b17d"
  dark-focus: "#75c899"
  dialog-backdrop: "#00000099"
  dialog-shadow: "#00000024"
typography:
  page-title-mobile:
    fontFamily: "Geist, sans-serif"
    fontSize: "32px"
    fontWeight: 600
    lineHeight: 1.2
    letterSpacing: "-0.03em"
  page-title-desktop:
    fontFamily: "Geist, sans-serif"
    fontSize: "36px"
    fontWeight: 600
    lineHeight: 1.2
    letterSpacing: "-0.03em"
  section-title:
    fontFamily: "Geist, sans-serif"
    fontSize: "24px"
    fontWeight: 600
    lineHeight: 1.4
    letterSpacing: "-0.025em"
  subsection-title:
    fontFamily: "Geist, sans-serif"
    fontSize: "18px"
    fontWeight: 600
    lineHeight: 1.5
    letterSpacing: "-0.02em"
  body:
    fontFamily: "Geist, sans-serif"
    fontSize: "16px"
    fontWeight: 400
    lineHeight: 1.75
  body-compact:
    fontFamily: "Geist, sans-serif"
    fontSize: "14px"
    fontWeight: 400
    lineHeight: 1.75
  control-label:
    fontFamily: "Geist, sans-serif"
    fontSize: "14px"
    fontWeight: 500
    lineHeight: 1.75
  navigation-desktop:
    fontFamily: "Geist, sans-serif"
    fontSize: "13px"
    fontWeight: 400
    lineHeight: 1.75
  navigation-group:
    fontFamily: "Geist, sans-serif"
    fontSize: "13px"
    fontWeight: 500
    lineHeight: 1.75
  caption:
    fontFamily: "Geist, sans-serif"
    fontSize: "12px"
    fontWeight: 400
    lineHeight: 1.8
  small-label:
    fontFamily: "Geist, sans-serif"
    fontSize: "12px"
    fontWeight: 400
    lineHeight: 1.75
  micro-label:
    fontFamily: "Geist, sans-serif"
    fontSize: "11px"
    fontWeight: 400
    lineHeight: 1.75
  code-mobile:
    fontFamily: "Geist Mono, monospace"
    fontSize: "12px"
    fontWeight: 400
    lineHeight: 1.85
  code-desktop:
    fontFamily: "Geist Mono, monospace"
    fontSize: "13px"
    fontWeight: 400
    lineHeight: 1.85
  code-inline:
    fontFamily: "Geist Mono, monospace"
    fontSize: ".85em"
    fontWeight: 400
    lineHeight: 1.75
  numeric:
    fontFamily: "Geist Mono, monospace"
  chart-value:
    fontFamily: "Geist Mono, monospace"
    fontSize: "13px"
    fontWeight: 400
    lineHeight: 1.75
  chart-unit-mobile:
    fontFamily: "Geist Mono, monospace"
    fontSize: "10px"
    fontWeight: 400
    lineHeight: 1.75
  chart-unit-wide:
    fontFamily: "Geist Mono, monospace"
    fontSize: "11px"
    fontWeight: 400
    lineHeight: 1.75
  benchmark-selector-mobile:
    fontFamily: "Geist, sans-serif"
    fontSize: "14px"
    fontWeight: 400
    lineHeight: 1.5
  benchmark-selector-desktop:
    fontFamily: "Geist, sans-serif"
    fontSize: "13px"
    fontWeight: 400
    lineHeight: 1.5
  benchmark-summary:
    fontFamily: "Geist, sans-serif"
    fontSize: "36px"
    fontWeight: 600
    lineHeight: 1.3
    letterSpacing: "-0.03em"
rounded:
  legend: "2px"
  bar: "3px"
  inline-code: "4px"
  control: "6px"
  brand-image: "9px"
  toast: "8px"
  panel: "10px"
  search-dialog: "12px"
spacing:
  "4": "4px"
  "6": "6px"
  "8": "8px"
  "10": "10px"
  "12": "12px"
  "14": "14px"
  "16": "16px"
  "18": "18px"
  "20": "20px"
  "24": "24px"
  "28": "28px"
  "32": "32px"
  "40": "40px"
  "48": "48px"
components:
  button-primary:
    backgroundColor: "{colors.ink}"
    textColor: "{colors.bg}"
    typography: "{typography.control-label}"
    rounded: "{rounded.control}"
    padding: "0 16px"
  button-icon:
    textColor: "{colors.muted}"
    rounded: "{rounded.control}"
    padding: "0"
    height: "44px"
    width: "44px"
  navbar-icon:
    backgroundColor: "{colors.bg}"
    textColor: "{colors.muted}"
    rounded: "{rounded.control}"
    padding: "0"
    height: "32px"
    width: "32px"
  navbar-search:
    backgroundColor: "{colors.bg}"
    textColor: "{colors.muted}"
    rounded: "{rounded.control}"
    padding: "0"
    height: "32px"
    width: "32px"
  navbar-search-desktop:
    backgroundColor: "{colors.surface}"
    padding: "0 10px"
    width: "clamp(208px, 20vw, 240px)"
  button-icon-hover:
    backgroundColor: "{colors.subtle}"
    textColor: "{colors.ink}"
  button-copy:
    textColor: "{colors.muted}"
    typography: "{typography.small-label}"
    rounded: "{rounded.control}"
    padding: "0 8px"
    height: "28px"
  code-panel-heading:
    textColor: "{colors.muted}"
    typography: "{typography.small-label}"
    padding: "0 12px 0 16px"
    height: "44px"
  code-panel-heading-desktop:
    height: "40px"
  search-input:
    backgroundColor: "{colors.bg}"
    textColor: "{colors.ink}"
    typography: "{typography.body}"
    padding: "10px 0"
  navigation-link:
    textColor: "{colors.muted}"
    typography: "{typography.body-compact}"
    padding: "8px 0"
  navigation-link-active:
    textColor: "{colors.green}"
    typography: "{typography.control-label}"
  code-panel:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.ink}"
    rounded: "{rounded.panel}"
  note:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.ink}"
    rounded: "{rounded.panel}"
    padding: "18px 16px"
  benchmark-selector:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.ink}"
    typography: "{typography.benchmark-selector-mobile}"
    rounded: "{rounded.control}"
    padding: "0 12px"
    height: "44px"
  benchmark-selector-desktop:
    typography: "{typography.benchmark-selector-desktop}"
    padding: "0 10px"
    height: "36px"
  benchmark-selector-hover:
    backgroundColor: "{colors.subtle}"
  benchmark-selector-popup:
    backgroundColor: "{colors.bg}"
    textColor: "{colors.ink}"
    rounded: "{rounded.control}"
    padding: "4px"
  benchmark-selector-option:
    typography: "{typography.benchmark-selector-mobile}"
    rounded: "{rounded.bar}"
    padding: "6px 8px"
    height: "44px"
  benchmark-selector-option-desktop:
    typography: "{typography.benchmark-selector-desktop}"
    height: "32px"
  benchmark-selector-option-active:
    backgroundColor: "{colors.subtle}"
  chart-panel:
    backgroundColor: "{colors.bg}"
    textColor: "{colors.ink}"
    rounded: "{rounded.panel}"
    padding: "20px 14px"
  chart-panel-wide:
    padding: "24px"
---

# Design System: MinkExcel documentation

## Overview

**Creative North Star: "Read mode"**

The implemented direction follows the user-approved Vercel Geist styling and Next.js documentation reference: quiet neutral surfaces, compact headings, precise navigation, and a clear reading column. The page title and technical content lead; the preserved MinkExcel logo provides the brand anchor.

The same documentation remains available across phone and desktop layouts. Color, borders, spacing, and authentic Solar Linear SVGs explain state and structure with little decoration. The light and dark themes share component geometry and typography.

**Key Characteristics:**
- Self-hosted Geist Sans and Geist Mono.
- Neutral reading surfaces with green links, active documentation navigation, and MinkExcel chart bars.
- Mobile-first layout with persistent navigation added as space permits.
- Flat bordered panels, visible numeric chart values, and native disclosure/dialog behavior.

Source of truth: `styles.css` defines appearance; `build.mjs`, `content.mjs`, `app.js`, `charts.mjs`, `numbers.mjs`, and `icons.mjs` define component structure and behavior. The frontmatter records implemented reusable values; the sidecar contains extensions and previews. Extracted role names describe current usage rather than introducing new CSS variables.

## Colors

A neutral palette supports long reading sessions, while the green accent identifies links, active documentation entries, selected text, focus outlines, and the MinkExcel benchmark series.

### Primary

`green` is the brand and interaction accent; `accent` is its pale selection background. `focus` matches the accent color in each theme. These are distinct semantic CSS properties even where values coincide.

### Neutral

`bg` is the page and dialog background, `surface` fills code panels and informational notes, and `subtle` fills inline code, hover states, and chart tracks. `ink` is primary text and the filled action background; `muted` is supporting text; `line` separates regions and outlines panels. `excel` is the contrasting slate used for the ExcelJS chart series.

`string`, `keyword`, and `number` distinguish syntax tokens; comments use `muted`. `dialog-backdrop` dims the page behind the search dialog and navigation sheet with a (60%) black tint, and `dialog-shadow` supplies the search dialog's soft shadow. Every `dark-*` token is the corresponding explicit or system-selected dark-theme value. The sidecar's synthesized tonal strips are preview aids, not additional implementation colors.

**The Semantic Color Rule.** Keep the green and slate chart series paired with visible names and numeric values; preserve each theme's semantic assignments.

## Typography

**Heading and Body Font:** Geist, with sans-serif fallback.
**Code and Numeric Font:** Geist Mono, with monospace fallback.

Both variable fonts are self-hosted with weights from 100 through 900 and `font-display: swap`; the observed interface uses regular (400), medium (500), and semibold (600). Retain the supplied font licenses.

### Hierarchy

- **Page titles:** the mobile and desktop title roles; balanced wrapping, semibold weight, tight tracking, and compact line height.
- **Section and subsection titles:** the two subordinate heading roles. Keep their hierarchy below the page title.
- **Body:** the body role applies to paragraphs and introductory descriptions. Supporting copy and mobile navigation use the compact role; controls use medium weight.
- **Navigation and labels:** desktop navigation/group labels use the named navigation roles; captions, small labels, and micro labels cover secondary metadata. Mobile navigation expands targets and text rather than shrinking body copy.
- **Code:** preformatted blocks use the mobile code role, switching to the desktop code role at the sidebar breakpoint. Inline code is relative to its surrounding text. Keep code horizontally scrollable within its panel.
- **Numeric text:** the numeric role applies Geist Mono with tabular numerals to visible numeric literals in prose, tables, versions, chart titles/captions, results, and enhanced selector labels. Size, weight, line height, and tracking come from the surrounding text. Axis ticks and the numeric-only native data-row selector also use mono.
- **Chart values:** the mono value role uses tabular numerals and separate small units. The benchmark summary role is a data emphasis, not a marketing heading.

**The Reading Type Rule.** Preserve the compact title scale and the body size across viewports; use natural wrapping rather than forced title line breaks.

**The Numeric Type Rule.** Use Geist Mono with tabular numerals for visible numeric text and values. Keep surrounding words and units in their existing font; existing whole chart values retain their mono units. Preserve identifiers such as CRC32, M4, and A1, and leave code rendering intact.

## Layout

The base layout is one column and is designed to fit a 320px viewport. Main content starts with (20px) horizontal gutters and (24px) top padding; the header and mobile navigation strip use (16px) gutters. The sticky top bar is (64px) tall, with a separate mobile strip at least (48px) tall. Sections begin (40px) apart on mobile.

At (540px), main padding becomes (32px), benchmark controls form three `minmax(0, 160px)` columns with (12px) gaps, paired code examples form two columns, and chart padding and labels expand. At (900px), the page shell adds a (240px) persistent sidebar, the mobile strip disappears, main padding becomes (40px 40px 28px), and the main column is capped at (850px) including padding. Page titles and code text switch to their desktop roles; section spacing becomes (48px). At (1280px), a (200px) right section list appears and main horizontal padding becomes (48px). The overall shell is capped at (1440px).

Mobile page sections use native `details`; the complete documentation navigation uses a native modal drawer with a no-JavaScript disclosure fallback. API rows place the signature above its description below (900px), then return to a two-column table. Comparison/result tables scroll inside their bordered regions. HTML bar charts reserve labels and values beside a flexible track; values remain visible on phones. Mobile navigation rows, dialog close actions, and benchmark controls retain targets at least (44px) tall; the persistent desktop navigation uses denser (36px) and nested (32px) rows. By explicit user choice, top navbar actions retain desktop density across viewports: (32px) controls, (16px) icons, and (6px) action gaps.

**The Contained Overflow Rule.** Keep the document within the viewport; isolate wide code and comparison data in their own scrolling regions.

## Elevation & Depth

The reading surface stays flat. Thin borders and the `surface`/`subtle` backgrounds separate content without raised cards. The search modal has the deepest elevation: its downward shadow is `0 16px 48px #00000024`, with the native dialog backdrop. Search and the navigation sheet share a (6px) backdrop blur through `backdrop-filter` and its `-webkit-` prefix in both themes; the foreground panels remain crisp. The navigation drawer uses the backdrop and a right border rather than a shadow. Search input focus uses `inset 0 -2px var(--focus)` on its enclosing row. Benchmark option popups use a smaller `0 4px 12px #00000026` shadow over the page background, separating the open list from the chart controls.

General keyboard focus is a (2px) accent outline offset by (4px). Benchmark selector triggers use the same accent at (2px) with a (2px) offset; forced-colors mode uses `Highlight` for this outline. Copy controls use a (2px) offset to fit within compact headers. Code scroll regions and search-result links use inset outlines with a (-2px) offset; code regions follow the panel’s (9px) inner bottom corners. Search input owns its row underline, while the close button keeps its independent outline. Forced-colors mode replaces the search underline with a (2px) inset `Highlight` outline on the row. Only when reduced motion is not requested, scrolling becomes smooth, the drawer enters from (-20px) over (.18s), and search fades as described below; both dialog effects use `cubic-bezier(.16, 1, .3, 1)`. No ornamental motion is required for reading, navigation, or charts.

## Shapes

The control, panel, and search-dialog radius roles distinguish small actions, code/data containers, and the modal search surface. Inline code has its own small radius; chart bars and legend swatches use tighter corners. The logo image and status toast retain their implemented radii. Only the small status dot is circular. Borders are (1px) throughout content panels and navigation separators.

## Components

### Buttons and icon actions

The primary text action uses reversed `ink`/`bg`, medium compact text, a control radius, and (16px) horizontal padding. Hover reduces opacity to (.85). Icon actions use muted icons on transparent backgrounds; hover adds the subtle surface and primary text color. Search, theme, copy, and drawer actions use authentic inline Solar Linear SVGs. The GitHub repository link uses the official Primer Octicons GitHub brand mark as the brand-specific exception. General icons are (20px), with (14–18px) variants for breadcrumbs, metadata, and navigation.

### Navbar actions

The top navbar brand pairs a (36px) green MinkExcel icon with two lines of text. The (20px) wordmark uses weight 650 for “Mink” and 400 for “Excel”; the muted (12px) “Documentation” subtitle sits (3px) below it. The icon has a (9px) radius and (10px) spacing beside the text. This reference-inspired lockup replaces the separate Documentation link and divider while preserving the (64px) navbar height across viewports.

Search, GitHub, theme, and mobile menu controls share a control radius and neutral (1px) border. Across all viewports, controls use a (32px) height, (16px) icons, (12px) labels, and (6px) action gaps; theme and GitHub are (32px) squares with zero padding. This compact navbar geometry follows the user-requested desktop size on phones as well. The mobile menu is an icon-only button at the end of the top navbar, matching the (32px) actions. Its separate Menu/version row is removed; version remains in the desktop sidebar. The menu is hidden at (900px), where the persistent sidebar appears. Standalone dialog close controls keep their (44px) size. GitHub centers a (16px) official Primer Octicons mark using `currentColor`; its icon-only link has a “GitHub repository” accessible label and title. Hover uses the subtle surface and primary text; focus uses the general keyboard outline. GitHub remains a bordered control without a hover underline. Keep the search trigger aligned with its neighboring controls, without an extra outer margin.

Below (900px), search and theme are (32px) squares and GitHub is available in the navigation drawer. At the desktop breakpoint, search expands to `clamp(208px, 20vw, 240px)` with (10px) horizontal padding and (6px) internal gaps; its (12px) label truncates with an ellipsis as needed. Its separate Geist Mono shortcut keycap has a (20px) minimum height, (10px) text, (4px) horizontal padding, and the inline-code radius, displaying ⌘ K on Apple platforms and Ctrl K elsewhere.

### Search field and dialogs

Search combines a (16px) input with a magnifier and close action inside the bordered input row. Input focus is conveyed by the row underline, with an inset outline in forced-colors mode; focusing the close action does not also highlight the input row. Results use compact text, control radii, and a subtle hover/focus fill. The search dialog is at most (640px) wide, at most (75dvh) tall, and fits the viewport with (12px) outer gutters. The navigation drawer is at most (360px) wide and (100dvh) tall. Native modal behavior and explicit close handlers restore focus to the triggering control. The native search cancel decoration is hidden so the Solar close action is the sole visible close control.

Search and its backdrop fade in over (140ms) and out over (100ms) using `cubic-bezier(.16, 1, .3, 1)` only when reduced motion is not requested. Animate opacity alone, keeping the dialog's position, size, content, and shadow fixed. Native `display` and `overlay` transitions retain the closing surface for the fade; `@starting-style` supplies the transparent entry state. Pointer events are disabled while closed and enabled while open. With reduced motion requested, opening and closing remain immediate.

Register search controls and shortcuts before requesting the search index. Cmd/Ctrl+K opens the dialog, and / opens it when focus is outside an editable control. Show an announced loading or unavailable state while data is pending or fails; retain a typed query and rerender it when the index arrives. Escape closes search even when its input contains text and returns focus to the navbar search trigger.

Restore the saved theme and platform keycap synchronously before styles paint. The document canvas and body use the same theme background. Keep navbar controls in place while application modules load, and replace the native mobile menu disclosure in its fixed (32px) navbar slot only after drawer handlers are ready. The fallback disclosure opens below the navbar and scrolls within the available viewport. Preload both font families; load interactive chart code only on the benchmarks page.

### Navigation

The desktop sidebar begins with a static Documentation label aligned with the navigation group headings and an inline muted Geist Mono version; this compact row has no border, icon, chevron, or link behavior.

Grouped links are muted at rest, become primary text on hover, and turn green with medium weight when the page is current. API subsection entries are indented beside a thin border. The desktop right section list instead marks the active section with primary text and a primary border. Preserve accessible current-page/current-location attributes, the compact breadcrumb, and adjacent-page links.

### Code and informational panels

Code panels use the surface fill, panel radius, border, and a separated header with copy action. The header is at least (44px) high below (900px) and (40px) high at the desktop breakpoint. The copy button uses its compact height and horizontal padding tokens, a (44px) minimum width, (12px) text, and a (16px) Solar icon. On coarse pointers, an invisible area extends (8px) above and below the button to preserve a (44px) touch target. Hover and keyboard focus retain the existing neutral fill and accent outline. The preformatted area scrolls independently; highlighting uses the semantic syntax colors. Informational notes share the flat panel treatment with a Solar document icon and compact copy. Copy success/failure feedback appears in the status toast. Code and content remain present when JavaScript is unavailable.

**The Compact Copy Rule.** Keep the copy button visually compact while preserving its full coarse-pointer touch target; the header follows the viewport breakpoint.

### Tables and benchmark charts

Table wrappers use the panel border/radius and a surface-filled heading row. API references use the responsive signature/description layout. Benchmark charts are HTML bars with named series, visible exact values, units, zero-based tracks, and a labeled scale; chart values and axis ticks use Geist Mono with tabular numerals. Numeric portions of titles, captions, table cells, and enhanced option labels use the shared numeric role while adjacent words retain their existing font. Default recorded chart data renders statically, while loaded data enables the controls.

Benchmark selectors use the compact trigger and bordered option-panel pattern from [shadcn/ui Select](https://ui.shadcn.com/docs/components/select), implemented in vanilla JavaScript with the native Popover API. The control radius, neutral (1px) border, `surface` fill, and `subtle` hover fill keep the trigger aligned with the other documentation controls. Each muted (12px) label sits (6px) above its control. A muted, authentic Solar Linear down-chevron sits on the right at (16px) square, with `aria-hidden` and `pointer-events: none`.

On mobile, triggers have a (44px) minimum height, (14px) text, and (12px) horizontal padding. At (900px) and above, triggers use a (36px) minimum height, (13px) text, and (10px) horizontal padding. The three controls retain their (160px) column cap from (540px). The open panel uses `bg`, a (1px) `line` border, the control radius, and (4px) padding. Options use (6px 8px) padding, tight (3px) corners, and a `subtle` active fill; they are at least (44px) high with (14px) text on mobile, and (32px) high with (13px) text at the desktop breakpoint. A right-aligned (16px) Solar Linear check identifies the committed selection. Forced-colors mode uses `Highlight`/`HighlightText` for the active option.

The trigger is a select-only combobox with an associated listbox. DOM focus stays on the trigger; expanded, active-descendant, and selected attributes convey state. Arrow, Home/End, Page Up/Down, and typeahead move the active option. Enter, Space, or Tab commit it; Escape cancels. Native auto-popover dismissal closes the panel on outside interaction. The panel uses fixed viewport positioning, flips above when needed, clamps within viewport gutters, and repositions on scrolling or resizing.

The hidden native select remains the value source and dispatches its existing change event to update charts. Browsers without the Popover API retain the visible native select, restoring native appearance and hiding its decorative SVG in forced-colors mode. Without JavaScript, the recorded default charts and result tables remain available. Runtime, operation, and data-row options and the recorded data are unchanged.

## Do's and Don'ts

### Do:
- **Do** preserve the approved Geist/Next.js direction and the MinkExcel logo.
- **Do** use the theme's semantic colors, authentic Solar Linear utility SVGs, and the official GitHub brand mark for the repository link.
- **Do** retain readable body copy, compact heading hierarchy, and mobile-first component structure.
- **Do** keep chart values visible and wide code/tables contained within their regions.
- **Do** preserve keyboard focus, native navigation fallbacks, static chart data, theme preferences, and reduced-motion behavior.

### Don't:
- **Don't** add oversized marketing titles, forced title line breaks, or ornamental background grids to the documentation.
- **Don't** replace the requested Solar utility icons with Unicode symbols or unrelated icon sets; keep the official GitHub brand mark as the repository-link exception.
- **Don't** imply a global reading measure or spacing value that the CSS does not implement.
- **Don't** apply synthesized sidecar tonal ramps as production tokens without an explicit design change.
