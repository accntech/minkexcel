const escape = text => String(text).replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]));

// Keep identifiers such as CRC32, M4 and A1 intact; format numeric literals only.
const numbers = /(?<![\p{L}\p{N}_&#])\d+(?:[.,]\d+)*(?![\p{L}\p{N}_])/gu;
const wrap = text => text.replace(numbers, value => `<span class="numeric">${value}</span>`);

export const numericText = text => wrap(escape(text));

// These fragments come from our own static templates. Preserve code, SVGs,
// native option labels and HTML entities while decorating visible text nodes.
export function numericMarkup(markup) {
  const protectedTags = new Set(['code', 'pre', 'svg', 'option', 'script', 'style']);
  let depth = 0;
  return markup.replace(/<span class="numeric">[^<]*<\/span>|<[^>]*>|[^<]+/g, token => {
    if (token.startsWith('<')) {
      const tag = token.match(/^<\/?([\w-]+)/)?.[1]?.toLowerCase();
      if (protectedTags.has(tag)) depth += token.startsWith('</') ? -1 : 1;
      return token;
    }
    return depth ? token : token.replace(/&(?:#\d+|#x[\da-f]+|\w+);|[^&]+|&/gi, part => part.startsWith('&') ? part : wrap(part));
  });
}
