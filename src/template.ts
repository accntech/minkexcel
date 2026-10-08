import {
  XlsxError,
  cellAddress,
  parseAddress,
  type Workbook,
  type Worksheet,
  type CellValue,
} from "./model.js";
import { attributes, child, excelText, xml, type XmlNode } from "./xml.js";

type TemplatePart = { path: string; xml: string; document: XmlNode };
export type TemplateRow = { node: XmlNode; cells: Map<number, XmlNode> };
export type TemplateSheet = TemplatePart & {
  sheet: Worksheet;
  rows: Map<number, TemplateRow>;
  values: Map<string, string>;
  presentation: string;
};
type Template = {
  parts: Map<string, Uint8Array>;
  sheets: TemplateSheet[];
  date1904: boolean;
  dateStyles: boolean[];
  workbook: TemplatePart;
  core?: TemplatePart;
  metadata: string;
  calculation: string;
};

// Private package state: callers cannot accidentally overwrite the retained archive.
export const templates = /* @__PURE__ */ new WeakMap<Workbook, Template>();

export function valueKey(value: CellValue): string {
  if (
    (value instanceof Date && !Number.isFinite(value.getTime())) ||
    (typeof value === "number" && !Number.isFinite(value))
  )
    throw new XlsxError("Cell numbers and dates must be finite.");
  return JSON.stringify(value);
}

function metadataValues(book: Workbook) {
  return {
    creator: book.creator,
    lastModifiedBy: book.lastModifiedBy ?? book.creator,
    created: book.created.toISOString(),
    modified: book.modified.toISOString(),
  };
}
export const metadataKey = (book: Workbook): string =>
  JSON.stringify(metadataValues(book));

function meaningful(value: unknown): boolean {
  if (value === undefined) return false;
  if (value && typeof value === "object")
    return Object.values(value).some(meaningful);
  return true;
}

export function presentationKey(sheet: Worksheet): string {
  const columns = [...sheet.columnDefinitions]
    .filter(
      ([, column]) => column.width !== undefined || meaningful(column.style),
    )
    .map(([index, column]) => [index, column.width, column.style]);
  const rows = [...sheet.rows]
    .filter(
      ([, row]) =>
        row.height !== undefined ||
        meaningful(row.style) ||
        [...row.cells.values()].some((cell) => meaningful(cell.style)),
    )
    .map(([index, row]) => [
      index,
      row.height,
      row.style,
      [...row.cells]
        .filter(([, cell]) => meaningful(cell.style))
        .map(([column, cell]) => [column, cell.style]),
    ]);
  return JSON.stringify({
    name: sheet.name,
    views: sheet.views,
    pageSetup: sheet.pageSetup,
    headerFooter: sheet.headerFooter,
    autoFilter: sheet.autoFilter,
    merges: sheet.merges,
    columns,
    rows,
  });
}

type Edit = { start: number; end: number; text: string };
type Address = ReturnType<typeof parseAddress>;
type ProtectedRange = {
  start: Address;
  end: Address;
  kind: "merge" | "formula";
};
type CellContext = {
  source: string;
  tag: (name: string) => string;
  columns: XmlNode[];
  template: Template;
};

function sourceOf(node: XmlNode): NonNullable<XmlNode["source"]> {
  if (!node.source)
    throw new XlsxError("Template XML source positions are missing.");
  return node.source;
}

function applyEdits(source: string, changes: Edit[]): string {
  // Apply from right to left so the original source offsets remain valid.
  for (const change of changes.sort((a, b) => b.start - a.start))
    source =
      source.slice(0, change.start) + change.text + source.slice(change.end);
  return source;
}

function raw(source: string, node: XmlNode): string {
  const span = sourceOf(node);
  return source.slice(span.start, span.end);
}

function replace(node: XmlNode, text: string): Edit {
  const { start, end } = sourceOf(node);
  return { start, end, text };
}

function qualified(node: XmlNode, local: string): string {
  const name = sourceOf(node).qualified;
  return name.slice(0, name.indexOf(":") + 1) + local;
}

function container(source: string, node: XmlNode, content: string): string {
  const span = sourceOf(node);
  const opening = source.slice(span.start, span.openEnd);
  if (opening.endsWith("/>"))
    return opening.slice(0, -2) + ">" + content + `</${span.qualified}>`;
  return opening + content + source.slice(span.closeStart, span.end);
}

function finite(value: number): string {
  if (!Number.isFinite(value))
    throw new XlsxError("Cell numbers must be finite.");
  return String(value);
}

function scalar(value: string | number | boolean | { error: string }): {
  type?: string;
  text: string;
} {
  if (typeof value === "number") return { text: finite(value) };
  if (typeof value === "boolean")
    return { type: "b", text: String(Number(value)) };
  if (typeof value === "string") return { type: "str", text: excelText(value) };
  return { type: "e", text: xml(value.error) };
}

function dateSerial(
  value: Date,
  style: string | undefined,
  template: Template,
): string {
  if (!template.dateStyles[Number(style ?? 0)])
    throw new XlsxError(
      "Template Date values require an existing date-formatted cell or column.",
    );
  const epoch = template.date1904
    ? Date.UTC(1904, 0, 1)
    : Date.UTC(1899, 11, 31);
  const days = (value.getTime() - epoch) / 86_400_000;
  return finite(!template.date1904 && days >= 60 ? days + 1 : days);
}

function cellContent(
  value: CellValue,
  attrs: Record<string, string | undefined>,
  context: CellContext,
): string {
  const { tag, template } = context;
  const v = (text: string) => `<${tag("v")}>${text}</${tag("v")}>`;
  if (value === null) return "";
  if (typeof value === "string") {
    attrs.t = "inlineStr";
    return `<${tag("is")}><${tag("t")} xml:space="preserve">${excelText(value)}</${tag("t")}></${tag("is")}>`;
  }
  if (value instanceof Date) return v(dateSerial(value, attrs.s, template));
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    (typeof value === "object" && "error" in value)
  ) {
    const result = scalar(value);
    attrs.t = result.type;
    return v(result.text);
  }
  if (typeof value === "object" && "formula" in value) {
    const formula = `<${tag("f")}>${xml(value.formula)}</${tag("f")}>`;
    if (value.result === undefined) return formula;
    const result = scalar(value.result);
    attrs.t = result.type;
    return formula + v(result.text);
  }
  throw new XlsxError("Unsupported cell value.");
}

function groupedFormula(node: XmlNode): XmlNode | undefined {
  const formula = child(node, "f");
  return formula?.attributes.t === "shared" || formula?.attributes.t === "array"
    ? formula
    : undefined;
}

function templateCell(
  value: CellValue,
  address: string,
  original: XmlNode | undefined,
  column: number,
  context: CellContext,
): string {
  if (original && groupedFormula(original))
    throw new XlsxError(
      "Template shared and array formula cells cannot be edited.",
    );
  const attrs: Record<string, string | undefined> = {
    ...original?.attributes,
    r: address,
  };
  delete attrs.t;
  if (!original) {
    const definition = context.columns.find(
      (node) =>
        Number(node.attributes.min) <= column &&
        Number(node.attributes.max) >= column,
    );
    if (definition?.attributes.style) attrs.s = definition.attributes.style;
  }
  const content = cellContent(value, attrs, context);
  // Keep cell extensions and metadata even when replacing its value.
  const extras = (original?.children ?? [])
    .filter((node) => !["f", "v", "is"].includes(node.name))
    .map((node) => raw(context.source, node))
    .join("");
  return `<${context.tag("c")}${attributes(attrs)}>${content}${extras}</${context.tag("c")}>`;
}

function protectedRanges(state: TemplateSheet): ProtectedRange[] {
  const ranges: ProtectedRange[] = [];
  const add = (ref: string, kind: ProtectedRange["kind"]) => {
    const [start, end = start] = ref.split(":").map(parseAddress);
    ranges.push({ start, end, kind });
  };
  for (const node of child(state.document, "mergeCells")?.children ?? [])
    add(node.attributes.ref, "merge");
  for (const row of state.rows.values()) {
    for (const cell of row.cells.values()) {
      const formula = groupedFormula(cell);
      if (formula?.attributes.ref) add(formula.attributes.ref, "formula");
    }
  }
  return ranges;
}

function assertEditable(
  row: number,
  column: number,
  ranges: ProtectedRange[],
): void {
  const protectedCell = ranges.some(({ start, end, kind }) => {
    if (
      row < start.row ||
      row > end.row ||
      column < start.column ||
      column > end.column
    )
      return false;
    return kind === "formula" || row !== start.row || column !== start.column;
  });
  if (protectedCell)
    throw new XlsxError(
      "Template merged members and shared/array formula ranges cannot be edited.",
    );
}

function sortedUnion(
  first: Iterable<number>,
  second: Iterable<number>,
): number[] {
  return [...new Set([...first, ...second])].sort((a, b) => a - b);
}

function templateRow(
  number: number,
  state: TemplateSheet,
  context: CellContext,
  ranges: ProtectedRange[],
): { text: string; changed: boolean } {
  const original = state.rows.get(number);
  const row = state.sheet.rows.get(number);
  const columns = sortedUnion(
    original?.cells.keys() ?? [],
    row?.cells.keys() ?? [],
  );
  let changed = false;
  const cells: string[] = [];
  for (const column of columns) {
    const originalCell = original?.cells.get(column);
    const address = cellAddress(number, column);
    const value = row?.cells.get(column)?.value ?? null;
    if (originalCell && valueKey(value) === state.values.get(address)) {
      cells.push(raw(state.xml, originalCell));
      continue;
    }
    if (!originalCell && value === null) continue;
    assertEditable(number, column, ranges);
    cells.push(templateCell(value, address, originalCell, column, context));
    changed = true;
  }
  if (original && !changed)
    return { text: raw(state.xml, original.node), changed };
  if (original) {
    const extras = original.node.children
      .filter((node) => node.name !== "c")
      .map((node) => raw(state.xml, node))
      .join("");
    return {
      text: container(state.xml, original.node, cells.join("") + extras),
      changed,
    };
  }
  const text = cells.length
    ? `<${context.tag("row")} r="${number}">${cells.join("")}</${context.tag("row")}>`
    : "";
  return { text, changed };
}

function dimensionEdit(state: TemplateSheet): Edit | undefined {
  const dimension = child(state.document, "dimension");
  if (!dimension) return undefined;
  let maximumRow = state.sheet.rowCount;
  let maximumColumn = state.sheet.columnCount;
  // Preserve styled blank extents rather than shrinking the template's dimensions.
  const last = dimension.attributes.ref?.split(":").at(-1);
  if (last && /^[A-Z]+\d+$/.test(last)) {
    const address = parseAddress(last);
    maximumColumn = Math.max(maximumColumn, address.column);
    maximumRow = Math.max(maximumRow, address.row);
  }
  const ref = `A1:${cellAddress(Math.max(maximumRow, 1), Math.max(maximumColumn, 1))}`;
  return replace(
    dimension,
    `<${sourceOf(dimension).qualified}${attributes({ ...dimension.attributes, ref })}/>`,
  );
}

function worksheetEdits(
  state: TemplateSheet,
  template: Template,
): string | undefined {
  if (presentationKey(state.sheet) !== state.presentation)
    throw new XlsxError(
      "Template mode supports cell value edits only; presentation changes are unsupported.",
    );
  const data = child(state.document, "sheetData");
  if (!data)
    throw new XlsxError("Template worksheet has no sheetData element.");
  const context: CellContext = {
    source: state.xml,
    tag: (name) => qualified(data, name),
    columns: child(state.document, "cols")?.children ?? [],
    template,
  };
  const ranges = protectedRanges(state);
  const rows: string[] = [];
  let changed = false;
  for (const number of sortedUnion(
    state.rows.keys(),
    state.sheet.rows.keys(),
  )) {
    const row = templateRow(number, state, context, ranges);
    rows.push(row.text);
    changed ||= row.changed;
  }
  if (!changed) return undefined;
  const changes = [replace(data, container(state.xml, data, rows.join("")))];
  const dimension = dimensionEdit(state);
  if (dimension) changes.push(dimension);
  return applyEdits(state.xml, changes);
}

function metadataEdits(
  book: Workbook,
  template: Template,
): { path: string; text: string } | undefined {
  const current = metadataValues(book);
  if (JSON.stringify(current) === template.metadata) return undefined;
  if (!template.core)
    throw new XlsxError("Template has no editable core metadata part.");
  const previous: ReturnType<typeof metadataValues> = JSON.parse(
    template.metadata,
  );
  const { xml: source, document } = template.core;
  const changes: Edit[] = [];
  for (const name of Object.keys(current) as Array<keyof typeof current>) {
    if (current[name] === previous[name]) continue;
    const node = child(document, name);
    if (!node)
      throw new XlsxError(`Template has no editable ${name} metadata field.`);
    changes.push(replace(node, container(source, node, xml(current[name]))));
  }
  return { path: template.core.path, text: applyEdits(source, changes) };
}

function recalculationEdit(part: TemplatePart): string {
  const { xml: source, document } = part;
  const calc = child(document, "calcPr");
  const text = `<${qualified(document, "calcPr")}${attributes({ ...calc?.attributes, fullCalcOnLoad: "1" })}/>`;
  if (calc) return applyEdits(source, [replace(calc, text)]);
  const extensions = child(document, "extLst");
  const offset = extensions
    ? sourceOf(extensions).start
    : sourceOf(document).closeStart;
  return applyEdits(source, [{ start: offset, end: offset, text }]);
}

export function templateArchive(
  book: Workbook,
): Map<string, string | Uint8Array> | undefined {
  const template = templates.get(book);
  if (!template) return undefined;
  if (
    book.worksheets.length !== template.sheets.length ||
    template.sheets.some(
      (state, index) => state.sheet !== book.worksheets[index],
    )
  )
    throw new XlsxError(
      "Template worksheets cannot be added, removed or reordered.",
    );
  if (JSON.stringify(book.calcProperties) !== template.calculation)
    throw new XlsxError("Template calculation properties cannot be edited.");
  const parts = new Map<string, string | Uint8Array>(template.parts);
  let changed = false;
  for (const state of template.sheets) {
    const output = worksheetEdits(state, template);
    if (output === undefined) continue;
    parts.set(state.path, output);
    changed = true;
  }
  const metadata = metadataEdits(book, template);
  if (metadata !== undefined) parts.set(metadata.path, metadata.text);
  if (changed)
    parts.set(template.workbook.path, recalculationEdit(template.workbook));
  return parts;
}
