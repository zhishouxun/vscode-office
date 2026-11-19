/** Visible text coordinates shared by search, bookmarks and caret restoration. */
export function visibleText(root) {
  const nodes = [];
  let text = '', previousBlock;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {
    const parent = node.parentElement;
    if (!node.textContent || parent.closest('[contenteditable="false"], script, style, .vditor-ir__marker, .vditor-ir__preview, .vditor-hint, .cm-gutters')) continue;
    if (!parent.getClientRects().length || getComputedStyle(parent).visibility === 'hidden') continue;
    const block = parent.closest('p, li, td, th, h1, h2, h3, h4, h5, h6, pre, .cm-line');
    if (nodes.length && block !== previousBlock) text += '\n';
    const start = text.length;
    text += node.textContent;
    nodes.push({ node, start, end: text.length });
    previousBlock = block;
  }
  return { text, nodes };
}

export function literalMatches(text, query) {
  if (!query) return [];
  const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'giu');
  const positions = [];
  let match;
  while ((match = pattern.exec(text))) {
    positions.push(match.index);
    pattern.lastIndex = match.index + (text.codePointAt(match.index) > 0xffff ? 2 : 1);
  }
  return positions;
}

export function textRange(snapshot, start, length = 0) {
  const first = snapshot.nodes.find(item => start >= item.start && start < item.end);
  const end = start + length;
  const last = length ? snapshot.nodes.find(item => end > item.start && end <= item.end) : first;
  if (!first || !last) return null;
  const range = document.createRange();
  range.setStart(first.node, start - first.start);
  range.setEnd(last.node, end - last.start);
  return range;
}

export function resolveAnchor(text, anchor) {
  if (!anchor || !Number.isFinite(anchor.offset)) return 0;
  const offset = Math.max(0, Math.min(Math.trunc(anchor.offset), Math.max(0, text.length - 1)));
  if (!anchor.quote || typeof anchor.quote !== 'string') return offset;
  const matchesContext = position => (!anchor.before || text.slice(Math.max(0, position - anchor.before.length), position) === anchor.before) &&
    (!anchor.after || text.slice(position + anchor.quote.length, position + anchor.quote.length + anchor.after.length) === anchor.after);
  if (text.slice(offset, offset + anchor.quote.length) === anchor.quote && matchesContext(offset)) return offset;
  const matches = literalMatches(text, anchor.quote);
  if (!matches.length) return offset;
  const contextual = matches.filter(matchesContext);
  return (contextual.length ? contextual : matches).reduce((best, position) => Math.abs(position - offset) < Math.abs(best - offset) ? position : best);
}
