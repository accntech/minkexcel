// Convert our controlled documentation templates, not arbitrary third-party HTML.
const decode = text => text.replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, entity) => {
  if (entity.startsWith('#')) return String.fromCodePoint(parseInt(entity.slice(entity[1].toLowerCase() === 'x' ? 2 : 1), entity[1].toLowerCase() === 'x' ? 16 : 10));
  return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[entity.toLowerCase()];
});
const plain = text => decode(text.replace(/<[^>]*>/g, ''));
const inline = html => decode(html
  .replace(/<code\b[^>]*>([\s\S]*?)<\/code>/g, (_, value) => '`' + value.replace(/<[^>]*>/g, '') + '`')
  .replace(/<a\b[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g, (_, url, label) => `[${label.replace(/<[^>]*>/g, '').trim()}](${url})`)
  .replace(/<(strong|b)\b[^>]*>([\s\S]*?)<\/\1>/g, '**$2**')
  .replace(/<br\s*\/?\s*>/g, ' ')
  .replace(/<[^>]*>/g, '')).replace(/\s+/g, ' ').trim();

export function documentationMarkdown(html) {
  const blocks = [];
  const save = block => { const key = `\u0000BLOCK${blocks.length}\u0000`; blocks.push(block); return `\n\n${key}\n\n`; };
  let result = html
    .replace(/<svg\b[^>]*>[\s\S]*?<\/svg>/g, '')
    .replace(/<button\b[^>]*>[\s\S]*?<\/button>/g, '')
    .replace(/<a class="anchor"[^>]*>[\s\S]*?<\/a>/g, '')
    .replace(/<pre\b([^>]*)>([\s\S]*?)<\/pre>/g, (_, attributes, body) => {
      const language = /Terminal code/.test(attributes) ? 'sh' : 'ts';
      return save('```' + language + '\n' + plain(body) + '\n```');
    })
    .replace(/<div class="(?:code-heading|hero-actions|chart-controls|legend|download-links)"[^>]*>[\s\S]*?<\/div>/g, '')
    .replace(/<table\b[^>]*>([\s\S]*?)<\/table>/g, (_, body) => {
      const rows = [...body.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/g)].map(row => [...row[1].matchAll(/<t[hd]\b[^>]*>([\s\S]*?)<\/t[hd]>/g)].map(cell => inline(cell[1]).replaceAll('|', '\\|')));
      if (!rows.length) return '';
      const line = cells => '| ' + cells.join(' | ') + ' |';
      return save([line(rows[0]), line(rows[0].map(() => '---')), ...rows.slice(1).map(line)].join('\n'));
    })
    .replace(/<h([2-4])\b[^>]*>([\s\S]*?)<\/h\1>/g, (_, level, body) => `\n\n${'#'.repeat(Number(level) + 1)} ${inline(body)}\n\n`)
    .replace(/<dt>([\s\S]*?)<\/dt>\s*<dd>([\s\S]*?)<\/dd>/g, (_, term, value) => `\n\n**${inline(term)}:** ${inline(value)}\n\n`)
    .replace(/<li\b[^>]*>([\s\S]*?)<\/li>/g, (_, body) => `\n- ${inline(body)}\n`)
    .replace(/<a\b[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g, (_, url, label) => `[${plain(label).trim()}](${decode(url)})`)
    .replace(/<code\b[^>]*>([\s\S]*?)<\/code>/g, (_, value) => '`' + value.replace(/<[^>]*>/g, '') + '`')
    .replace(/<(strong|b)\b[^>]*>([\s\S]*?)<\/\1>/g, '**$2**')
    .replace(/<\/?(?:p|div|section|aside|article|dl|dt|dd|noscript)\b[^>]*>/g, '\n\n')
    .replace(/<\/span>/g, ' ')
    .replace(/<[^>]*>/g, '');
  result = decode(result).split('\n').map(line => line.trim()).join('\n').replace(/\n{3,}/g, '\n\n').trim();
  return result.replace(/\u0000BLOCK(\d+)\u0000/g, (_, index) => blocks[Number(index)]);
}

export function llmsFullGuide({ introduction, documents, version, exports, limitFields, datasets }) {
  const split = introduction.indexOf('\n## Documentation');
  const overview = split < 0 ? introduction : introduction.slice(0, split);
  const references = split < 0 ? '' : introduction.slice(split);
  const declarations = `## Public exports and import-limit type\n\nImport public names from \`minkexcel\`. This is the package entry point; the relative re-export paths below describe internal modules.\n\n\`\`\`ts\n${exports.trim()}\n\`\`\`\n\nThe expanded public import-limit type is:\n\n\`\`\`ts\nexport type ReadLimits = Partial<{\n${limitFields.map(field => `  ${field}: number;`).join('\n')}\n}>;\nexport type ReadOptions = ReadLimits & { preserveTemplate?: boolean };\n\`\`\``;
  const content = documents.map(page => {
    const body = page.slug === 'benchmarks' ? page.body
      .replace('Switch the runtime, operation or row count to compare the same three workloads. Every bar begins at zero. Lower is better.', 'Compare numeric, text and mixed workloads at each row count using the complete tables below. Lower processing time is better.')
      .replace(/<noscript>[\s\S]*?<\/noscript>/g, '')
      .replace(/<div class="chart-top">[\s\S]*?<\/div><\/div>/g, '')
      .replace(/<p id="table-caption">[\s\S]*?<\/p>/g, '<p>Complete recorded timings and file sizes for both runtimes follow below.</p>')
      : page.body;
    return `## ${page.label}\n\n${page.description}\n\n${documentationMarkdown(body)}`;
  }).join('\n\n');
  const measurements = datasets.map(data => {
    const metadata = data.metadata;
    const heading = `## Recorded results: ${metadata.runtime}\n\nMeasured ${metadata.date}; ${metadata.cpu}; ${metadata.platform}; ExcelJS ${metadata.exceljs}. Each time is the median of ${metadata.iterations} samples after ${metadata.warmups} warmup. Workloads have ${metadata.columns} columns plus a header. AbortSignal supplied: ${metadata.abortSignal}. Times below are rounded to two decimal places; file sizes are exact compressed bytes.\n\n`;
    const headers = '| Workload | Data rows | MinkExcel export ms | ExcelJS export ms | MinkExcel import ms | ExcelJS import ms | MinkExcel bytes | ExcelJS bytes |\n| --- | --- | --- | --- | --- | --- | --- | --- |';
    const rows = data.rows.map(row => '| ' + [row.workload, row.rows, ...['minkexcelExportMs', 'exceljsExportMs', 'minkexcelImportMs', 'exceljsImportMs'].map(key => row[key].toFixed(2)), row.minkexcelBytes, row.exceljsBytes].join(' | ') + ' |').join('\n');
    return heading + headers + '\n' + rows;
  }).join('\n\n');
  return `${overview.trim()}\n\nSelf-contained documentation for MinkExcel ${version}. This file includes the six published documentation pages, their complete examples and API tables, public types, import defaults, error handling, migration guidance, and all recorded benchmark medians and file sizes. Other links are optional references; fetching them is not necessary to use this guide. Generated from the website content, package declarations and recorded datasets.\n\n${declarations}\n\n${content}\n\n${measurements}\n\nOptional references to the same documentation and source material follow.\n\n${references.trim()}\n`;
}
