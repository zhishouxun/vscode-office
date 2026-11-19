/** MIT. Local large-document support for Office Viewer. Pure source model; no DOM. */
export const LARGE_DOCUMENT_CHARS = 48000;
export const PAGE_CHARS = 48000;
export const PAGE_LINES = 320;
const linesOf = text => [...text.matchAll(/[^\n]*\n|[^\n]+$/g)].map(m => ({ text: m[0], start: m.index, end: m.index + m[0].length }));
const trim = line => line.trim();
export function isTableSeparator(line) {
  const cells = trim(line).replace(/^\|/, '').replace(/\|$/, '').split('|');
  return cells.length > 0 && cells.every(c => /^:?-{3,}:?$/.test(c.trim()));
}
function tableRow(line) { return trim(line).includes('|') && trim(line) !== ''; }
function scanLine(state, line, next) {
  const s = trim(line);
  const fence = s.match(/^(`{3,}|~{3,})/);
  if (fence) {
    if (!state.fence) state.fence = { char: fence[1][0], size: fence[1].length };
    else if (fence[1][0] === state.fence.char && fence[1].length >= state.fence.size && new RegExp(`^${state.fence.char === '`' ? '`' : '~'}+$`).test(s)) state.fence = null;
    state.header = ''; return;
  }
  if (state.fence) return;
  if (tableRow(line) && next !== undefined && isTableSeparator(next)) state.header = line + next;
  else if (state.header && !tableRow(line)) state.header = '';
}
export function splitDocument(text, maxChars = PAGE_CHARS, maxLines = PAGE_LINES) {
  const lines = linesOf(text), spans = [];
  const state = { fence: null, header: '' };
  let start = 0, count = 0, frontmatter = lines.length > 0 && trim(lines[0].text) === '---';
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const separator = i > 0 && isTableSeparator(line.text) && tableRow(lines[i - 1].text);
    if (line.start > start && !state.fence && !frontmatter && !separator && (line.start - start >= maxChars || count >= maxLines)) {
      spans.push({ start, end: line.start }); start = line.start; count = 0;
    }
    scanLine(state, line.text, lines[i + 1]?.text);
    if (frontmatter && i > 0 && /^(---|\.\.\.)$/.test(trim(line.text))) frontmatter = false;
    count++;
  }
  spans.push({ start, end: text.length });
  return spans;
}
function stripSyntheticHeader(rendered, expectedPrefix) {
  const lines = linesOf(rendered);
  // Vditor may insert a blank paragraph before the table.
  let i = 0;
  while (i < lines.length && !trim(lines[i].text)) i++;
  if (!lines[i] || !lines[i + 1] || !tableRow(lines[i].text) || !isTableSeparator(lines[i + 1].text)) {
    throw new Error('此页的重复表头不能删除。请到表格第一页修改表头。');
  }
  const expected = linesOf(expectedPrefix);
  const normalizeHeader = text => text.replace(/\s+/g, '').replace(/^\|/, '').replace(/\|$/, '');
  if (normalizeHeader(lines[i].text) !== normalizeHeader(expected[0].text)) {
    throw new Error('此页的重复表头不能修改。请到表格第一页修改表头。');
  }
  return rendered.slice(lines[i + 1].end);
}
export class LargeDocument {
  constructor(text, options = {}) {
    this.maxChars = options.maxChars || PAGE_CHARS;
    this.maxLines = options.maxLines || PAGE_LINES;
    this.enabled = options.force ?? (text.length >= LARGE_DOCUMENT_CHARS || linesOf(text).length > 400);
    this.reset(text);
  }
  reset(text) {
    this.text = text;
    this.spans = this.enabled ? splitDocument(text, this.maxChars, this.maxLines) : [{ start: 0, end: text.length }];
    this.index = Math.min(this.index || 0, this.spans.length - 1);
    this.prefix = '';
    this.invalidate();
  }
  invalidate() {
    this.pageCache = new Map();
    this.tablePrefixes = null;
    this.searchCache = null;
  }
  preparePage(index) {
    if (!Number.isInteger(index) || index < 0 || index >= this.spans.length) return null;
    if (this.pageCache.has(index)) {
      const cached = this.pageCache.get(index);
      this.pageCache.delete(index); this.pageCache.set(index, cached);
      return cached;
    }
    if (index > 0 && !this.tablePrefixes) {
      // Scan once per document revision, rather than scanning the whole file on every turn.
      const lines = linesOf(this.text), state = { fence: null, header: '' };
      this.tablePrefixes = new Array(this.spans.length).fill('');
      let page = 1;
      for (let i = 0; i < lines.length && page < this.spans.length; i++) {
        if (lines[i].start === this.spans[page].start) {
          this.tablePrefixes[page++] = !state.fence && state.header && tableRow(lines[i].text) && !isTableSeparator(lines[i].text) ? state.header : '';
        }
        scanLine(state, lines[i].text, lines[i + 1]?.text);
      }
    }
    // The first page never needs a synthetic header; leave indexing to idle preloading.
    const span = this.spans[index], prefix = index > 0 ? this.tablePrefixes[index] : '';
    const cached = { prefix, value: prefix + this.text.slice(span.start, span.end) };
    this.pageCache.set(index, cached);
    // Keep at most the current and two adjacent prepared pages, never detached editor DOM.
    if (this.pageCache.size > 3) this.pageCache.delete(this.pageCache.keys().next().value);
    return cached;
  }
  setPageSize(maxChars, maxLines) {
    if (!Number.isInteger(maxChars) || maxChars < 1 || !Number.isInteger(maxLines) || maxLines < 1) {
      throw new RangeError('每页容量必须为正整数');
    }
    const offset = this.spans[this.index].start;
    this.maxChars = maxChars;
    this.maxLines = maxLines;
    this.reset(this.text);
    return this.indexAt(offset);
  }
  page(index = this.index) {
    this.index = Math.max(0, Math.min(index, this.spans.length - 1));
    const cached = this.preparePage(this.index);
    this.prefix = cached.prefix;
    return cached.value;
  }
  commit(rendered) {
    const span = this.spans[this.index];
    let replacement = this.prefix ? stripSyntheticHeader(rendered, this.prefix) : rendered;
    const original = this.text.slice(span.start, span.end);
    const eol = original.includes('\r\n') ? '\r\n' : '\n';
    // Editors normalize EOLs; retain the source document's convention and page boundary.
    replacement = replacement.replace(/\r?\n/g, eol);
    if (span.end < this.text.length && original.endsWith('\n') && !replacement.endsWith('\n')) replacement += eol;
    const delta = replacement.length - (span.end - span.start);
    this.text = this.text.slice(0, span.start) + replacement + this.text.slice(span.end);
    span.end += delta;
    for (let i = this.index + 1; i < this.spans.length; i++) { this.spans[i].start += delta; this.spans[i].end += delta; }
    this.invalidate();
    return this.text;
  }
  indexAt(offset) {
    const index = this.spans.findIndex(s => offset >= s.start && offset < s.end);
    return index < 0 ? this.spans.length - 1 : index;
  }
  matches(query) {
    if (!query) return [];
    if (this.searchCache?.query === query) return this.searchCache.positions;
    const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'giu');
    const positions = [];
    let match;
    while ((match = pattern.exec(this.text))) {
      positions.push(match.index);
      pattern.lastIndex = match.index + (this.text.codePointAt(match.index) > 0xffff ? 2 : 1);
    }
    this.searchCache = { query, positions };
    return positions;
  }
  find(query, from = 0, backwards = false) {
    const positions = this.matches(query);
    if (!positions.length) return -1;
    const start = Math.max(0, Math.min(from, this.text.length));
    let low = 0, high = positions.length;
    while (low < high) { const mid = (low + high) >>> 1; if (positions[mid] < start) low = mid + 1; else high = mid; }
    return backwards ? positions[(low - 1 + positions.length) % positions.length] : positions[low % positions.length];
  }
}
