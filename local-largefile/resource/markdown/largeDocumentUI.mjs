import { LargeDocument } from './largeDocument.mjs';
import { visibleText, literalMatches, textRange, resolveAnchor } from './documentNavigation.mjs';

export function createLargeDocumentUI(content, onChange, options = {}) {
  const pageSizes = [160, 320, 640];
  const loadPreference = key => { try { return localStorage.getItem(key); } catch { return null; } };
  const savePreference = (key, value) => { try { localStorage.setItem(key, String(value)); } catch { /* Storage may be unavailable. */ } };
  const savedSize = Number(loadPreference('office.largeDocument.pageLines'));
  const pageLines = pageSizes.includes(savedSize) ? savedSize : 320;
  const model = new LargeDocument(content, { maxLines: pageLines, maxChars: pageLines * 150 });
  const positionKey = options.documentId ? `office.largeDocument.position:${options.documentId}` : null;
  let initialBookmark;
  try { initialBookmark = positionKey ? JSON.parse(loadPreference(positionKey)) : null; } catch { initialBookmark = null; }
  if (model.enabled && initialBookmark?.anchor) model.page(model.indexAt(resolveAnchor(content, initialBookmark.anchor)));
  let restoring = !!initialBookmark, positionTimer, preloadTask;
  let autoPaging = loadPreference('office.largeDocument.autoPaging') !== 'false';
  let editor, readPage, baseline, ready = false, rendering = false, searchOffset = -1;
  let bar, select, status, query, previous, next;
  let wheelDistance = 0, wheelDirection = 0, wheelBlocked = false, wheelIdle, wheelBlockedUntil = 0, composing = false;
  const activeRoot = () => editor?.vditor?.currentMode === 'ir' ? editor?.vditor?.ir?.element : editor?.vditor?.wysiwyg?.element;
  const scrollContainer = root => {
    for (let node = root; node && node !== document.body; node = node.parentElement) {
      if (/(auto|scroll)/.test(getComputedStyle(node).overflowY) && node.scrollHeight > node.clientHeight + 2) return node;
    }
    return root;
  };
  const resetWheel = () => { wheelDistance = 0; wheelDirection = 0; };
  const reportError = error => { if (status) { status.textContent = error.message; status.setAttribute('role', 'alert'); } };
  const refresh = () => {
    if (!bar) return;
    if (select.options.length !== model.spans.length) {
      select.replaceChildren();
      for (let i = 0; i < model.spans.length; i++) {
        const option = document.createElement('option'); option.value = String(i); option.textContent = `第 ${i + 1} / ${model.spans.length} 页`; select.append(option);
      }
    }
    select.value = String(model.index);
    previous.disabled = model.index === 0; next.disabled = model.index === model.spans.length - 1;
    status.textContent = autoPaging ? '页底继续下滚翻页 · 页顶上滚返回 · 保存为一个文件' : '手动翻页 · 保存为一个文件'; status.removeAttribute('role');
  };
  const protectTableHeader = () => {
    if (!model.prefix) return;
    const root = editor?.vditor?.currentMode === 'ir' ? editor?.vditor?.ir?.element : editor?.vditor?.wysiwyg?.element;
    const header = root?.querySelector('table thead');
    if (header) { header.setAttribute('contenteditable', 'false'); header.title = '重复表头；请到表格第一页修改'; }
  };
  const commit = rendered => {
    if (!ready || rendering || rendered === baseline) return model.text;
    try { model.commit(rendered); baseline = rendered; searchOffset = -1; warmNeighbors(); return model.text; }
    catch (error) { reportError(error); throw error; }
  };
  const read = () => commit(readPage?.() ?? '');
  const warmNeighbors = () => {
    if (preloadTask != null) {
      if (window.cancelIdleCallback) window.cancelIdleCallback(preloadTask); else clearTimeout(preloadTask);
    }
    const index = model.index, text = model.text;
    const prepare = () => {
      preloadTask = null;
      if (model.index !== index || model.text !== text) return;
      model.preparePage(index - 1); model.preparePage(index + 1);
    };
    preloadTask = window.requestIdleCallback ? window.requestIdleCallback(prepare, { timeout: 800 }) : setTimeout(prepare, 0);
  };
  const sourceAnchor = (quote, approximate) => {
    const span = model.spans[model.index];
    const matches = literalMatches(model.text.slice(span.start, span.end), quote).map(offset => offset + span.start);
    const offset = matches.length ? matches.reduce((best, position) => Math.abs(position - approximate) < Math.abs(best - approximate) ? position : best) : span.start;
    return { offset, quote: matches.length ? quote : '', before: model.text.slice(Math.max(0, offset - 32), offset), after: model.text.slice(offset + quote.length, offset + quote.length + 32) };
  };
  const capturePosition = () => {
    const root = activeRoot();
    if (!root) return null;
    const scroller = scrollContainer(root), rect = scroller.getBoundingClientRect(), snapshot = visibleText(root);
    const span = model.spans[model.index];
    const ratio = scroller.scrollTop / Math.max(1, scroller.scrollHeight - scroller.clientHeight);
    const makeAnchor = (item, offset = 0) => {
      const start = Math.max(0, Math.min(offset, item.node.textContent.length - 1));
      const quote = item.node.textContent.slice(start, start + 64);
      return sourceAnchor(quote, span.start + (span.end - span.start) * item.start / Math.max(1, snapshot.text.length));
    };
    let item, visibleOffset = 0;
    const point = document.caretRangeFromPoint?.(rect.left + Math.min(80, rect.width / 3), rect.top + 16);
    if (point) {
      item = snapshot.nodes.find(entry => entry.node === point.startContainer);
      if (item) visibleOffset = point.startOffset;
    }
    if (!item) item = snapshot.nodes.find(entry => {
      const range = textRange(snapshot, entry.start, Math.min(1, entry.end - entry.start));
      const box = range?.getBoundingClientRect(); return box && box.bottom > rect.top && box.top < rect.bottom;
    });
    const anchor = item ? makeAnchor(item, visibleOffset) : { offset: span.start, quote: '' };
    const range = item ? textRange(snapshot, item.start + Math.min(visibleOffset, item.end - item.start - 1), 1) : null;
    const bookmark = { anchor, ratio, topDelta: range ? range.getBoundingClientRect().top - rect.top : 0, caret: null };
    const selection = window.getSelection();
    const selected = selection?.isCollapsed && snapshot.nodes.find(entry => entry.node === selection.anchorNode);
    if (selected) {
      const quoteStart = Math.max(0, selection.anchorOffset - 24);
      bookmark.caret = { anchor: makeAnchor(selected, quoteStart), delta: selection.anchorOffset - quoteStart };
    }
    return bookmark;
  };
  const savePosition = () => {
    clearTimeout(positionTimer);
    if (!positionKey || !ready || !model.enabled || rendering || restoring) return;
    const bookmark = capturePosition();
    if (bookmark) savePreference(positionKey, JSON.stringify(bookmark));
  };
  const schedulePosition = () => {
    if (restoring || rendering || !ready) return;
    clearTimeout(positionTimer); positionTimer = setTimeout(savePosition, 180);
  };
  const rangeForAnchor = (snapshot, anchor, length = 1) => {
    if (!anchor?.quote) return null;
    const absolute = resolveAnchor(model.text, anchor), span = model.spans[model.index];
    if (absolute < span.start || absolute >= span.end) return null;
    const sourceMatches = literalMatches(model.text.slice(span.start, span.end), anchor.quote);
    const ordinal = sourceMatches.indexOf(absolute - span.start);
    const renderedMatches = literalMatches(snapshot.text, anchor.quote);
    if (ordinal < 0 || sourceMatches.length !== renderedMatches.length) return null;
    return textRange(snapshot, renderedMatches[ordinal], length);
  };
  const restorePosition = (bookmark, focus = false) => {
    const root = activeRoot();
    if (!root || !bookmark) return;
    const snapshot = visibleText(root), scroller = scrollContainer(root);
    if (focus && bookmark.caret?.anchor) {
      const caret = rangeForAnchor(snapshot, bookmark.caret.anchor, bookmark.caret.delta);
      if (caret) {
        caret.collapse(false); root.focus({ preventScroll: true });
        const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(caret);
      }
    }
    const range = rangeForAnchor(snapshot, bookmark.anchor);
    if (range) scroller.scrollTop += range.getBoundingClientRect().top - scroller.getBoundingClientRect().top - (bookmark.topDelta || 0);
    else scroller.scrollTop = Math.max(0, Math.min(1, bookmark.ratio || 0)) * (scroller.scrollHeight - scroller.clientHeight);
    if (document.scrollingElement) document.scrollingElement.scrollTop = 0;
  };
  const render = (index, { edge = 'start', focus = false } = {}) => {
    rendering = true;
    try {
      const page = model.page(index);
      editor.setValue(page, true);
      baseline = readPage();
      protectTableHeader();
      const root = activeRoot();
      if (root && focus) {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
          acceptNode: node => node.parentElement?.closest('[contenteditable="false"]') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
        });
        let node, target;
        while ((node = walker.nextNode())) { target = node; if (edge === 'start') break; }
        root.focus({ preventScroll: true });
        if (target) {
          const range = document.createRange(); range.setStart(target, edge === 'start' ? 0 : target.textContent.length); range.collapse(true);
          const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(range);
        }
      }
      if (root) { const scroller = scrollContainer(root); scroller.scrollTop = edge === 'end' ? scroller.scrollHeight : 0; }
      if (document.scrollingElement) document.scrollingElement.scrollTop = 0;
      refresh(); warmNeighbors();
    } finally { rendering = false; schedulePosition(); }
  };
  const go = (index, options) => {
    if (!ready || index < 0 || index >= model.spans.length || index === model.index) return false;
    try { const text = read(); onChange(text); render(index, options); resetWheel(); return true; }
    catch (error) { reportError(error); return false; }
  };
  const canAutoPage = (event, direction) => {
    const root = activeRoot();
    if (!autoPaging || !ready || rendering || restoring || composing || !root?.contains(event.target) ||
        event.target.closest('input, textarea, select, button, dialog, .cm-editor, [contenteditable="false"]') ||
        event.ctrlKey || event.metaKey || event.altKey || event.shiftKey || event.isComposing || !window.getSelection()?.isCollapsed) return false;
    if (direction > 0 ? model.index === model.spans.length - 1 : model.index === 0) return false;
    const scroller = scrollContainer(root);
    // A code block or another nested scroll region must consume its own scroll first.
    for (let node = event.target; node && node !== scroller; node = node.parentElement) {
      if (node === root) continue;
      if (/(auto|scroll)/.test(getComputedStyle(node).overflowY) && node.scrollHeight > node.clientHeight + 2 &&
          (direction > 0 ? node.scrollTop + node.clientHeight < node.scrollHeight - 2 : node.scrollTop > 2)) return false;
    }
    return direction > 0 ? scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 2 : scroller.scrollTop <= 2;
  };
  const onWheel = event => {
    const scheduleIdle = () => {
      clearTimeout(wheelIdle);
      wheelIdle = setTimeout(() => { wheelBlocked = false; resetWheel(); }, Math.max(180, wheelBlockedUntil - performance.now()));
    };
    scheduleIdle();
    if (!event.deltaY || Math.abs(event.deltaX) > Math.abs(event.deltaY)) return;
    const direction = Math.sign(event.deltaY);
    if (!canAutoPage(event, direction)) { resetWheel(); return; }
    if (wheelBlocked) return;
    if (direction !== wheelDirection) resetWheel();
    wheelDirection = direction;
    const scale = event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? scrollContainer(activeRoot()).clientHeight : 1;
    wheelDistance += Math.abs(event.deltaY) * scale;
    if (wheelDistance < 100) return;
    event.preventDefault();
    wheelBlocked = true; // Ignore the remaining momentum until this gesture ends.
    go(model.index + direction, { edge: direction > 0 ? 'start' : 'end' });
    // Start the guard after rendering: expensive pages must not consume the idle window.
    wheelBlockedUntil = performance.now() + 500;
    scheduleIdle();
  };
  const onPageKey = event => {
    const direction = event.key === 'PageDown' ? 1 : event.key === 'PageUp' ? -1 : 0;
    if (!direction || event.repeat || !canAutoPage(event, direction)) return;
    if (go(model.index + direction, { edge: direction > 0 ? 'start' : 'end', focus: true })) {
      event.preventDefault(); event.stopPropagation(); wheelBlocked = true;
    }
  };
  const search = backwards => {
    if (!ready) return;
    try {
      const text = read(); onChange(text);
      const value = query.value;
      const from = searchOffset < 0 ? 0 : (backwards ? searchOffset : searchOffset + 1);
      const position = model.find(value, from, backwards);
      if (position < 0) { status.textContent = '未找到匹配内容'; return; }
      const matches = model.matches(value);
      searchOffset = position;
      const target = model.indexAt(position);
      if (target !== model.index) render(target);
      const span = model.spans[target], pageMatches = matches.filter(offset => offset >= span.start && offset < span.end);
      const ordinal = pageMatches.indexOf(position);
      const snapshot = visibleText(activeRoot()), renderedMatches = literalMatches(snapshot.text, value);
      const range = pageMatches.length === renderedMatches.length ? textRange(snapshot, renderedMatches[ordinal], value.length) : null;
      status.textContent = `第 ${matches.indexOf(position) + 1} / ${matches.length} 处 · 第 ${model.index + 1} 页`;
      if (range) {
        const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(range);
        const scroller = scrollContainer(activeRoot());
        scroller.scrollTop += range.getBoundingClientRect().top - scroller.getBoundingClientRect().top - scroller.clientHeight / 2;
        if (document.scrollingElement) document.scrollingElement.scrollTop = 0;
      } else {
        window.getSelection()?.removeAllRanges();
        status.textContent += ' · 此源码匹配无法在预览中准确高亮，可切换源码查看';
      }
      schedulePosition();
    } catch (error) { reportError(error); }
  };
  if (model.enabled) {
    document.body.classList.add('office-large-document');
    bar = document.createElement('nav'); bar.className = 'office-large-document-bar'; bar.setAttribute('aria-label', '大文件分页与全文搜索');
    const makeButton = (text, action, label) => { const b = document.createElement('button'); b.type = 'button'; b.textContent = text; b.setAttribute('aria-label', label || text); b.addEventListener('click', action); bar.append(b); return b; };
    previous = makeButton('上一页', () => go(model.index - 1));
    select = document.createElement('select'); select.setAttribute('aria-label', '选择页码'); select.addEventListener('change', () => go(Number(select.value))); bar.append(select);
    next = makeButton('下一页', () => go(model.index + 1));
    const sizeSelect = document.createElement('select'); sizeSelect.setAttribute('aria-label', '每页容量');
    for (const lines of pageSizes) {
      const option = document.createElement('option'); option.value = String(lines); option.textContent = `${lines} 行 / ${lines * 150 / 1000}k 字符`; sizeSelect.append(option);
    }
    sizeSelect.value = String(pageLines);
    sizeSelect.addEventListener('change', () => {
      try {
        const text = read(); onChange(text);
        const bookmark = capturePosition();
        const lines = Number(sizeSelect.value);
        model.setPageSize(lines * 150, lines);
        const index = model.indexAt(resolveAnchor(model.text, bookmark?.anchor));
        render(index); restorePosition(bookmark, true); resetWheel(); searchOffset = -1;
        savePreference('office.largeDocument.pageLines', lines);
      } catch (error) { reportError(error); }
    });
    bar.append(sizeSelect);
    const autoLabel = document.createElement('label'); autoLabel.className = 'office-large-document-auto';
    const autoInput = document.createElement('input'); autoInput.type = 'checkbox'; autoInput.checked = autoPaging;
    autoInput.addEventListener('change', () => { autoPaging = autoInput.checked; resetWheel(); savePreference('office.largeDocument.autoPaging', autoPaging); refresh(); });
    autoLabel.append(autoInput, '自动翻页'); bar.append(autoLabel);
    query = document.createElement('input'); query.type = 'search'; query.placeholder = '搜索整篇文件'; query.setAttribute('aria-label', '搜索整篇文件'); query.addEventListener('input', () => { searchOffset = -1; }); query.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); search(e.shiftKey); } }); bar.append(query);
    makeButton('查找', () => search(false));
    makeButton('源码', () => handler.emit('editInVSCode', true), '在 VS Code 中打开完整 Markdown 源码');
    status = document.createElement('span'); status.className = 'office-large-document-status'; status.setAttribute('aria-live', 'polite'); bar.append(status);
    document.body.prepend(bar); refresh();
    bar.setAttribute('aria-busy', 'true');
    bar.querySelectorAll('button, input, select').forEach(control => { control.disabled = true; });
  }
  return {
    initialValue: model.page(),
    enabled: model.enabled,
    read,
    input: commit,
    attach(instance, reader) {
      editor = instance; readPage = reader; baseline = reader(); ready = true; protectTableHeader();
      if (model.enabled) new MutationObserver(protectTableHeader).observe(document.getElementById('vditor'), { childList: true, subtree: true });
      if (model.enabled) {
        const host = document.getElementById('vditor');
        host.addEventListener('wheel', onWheel, { passive: false });
        host.addEventListener('keydown', onPageKey, true);
        host.addEventListener('compositionstart', () => { composing = true; });
        host.addEventListener('compositionend', () => { composing = false; });
        host.addEventListener('scroll', schedulePosition, true);
        host.addEventListener('keyup', schedulePosition);
        host.addEventListener('mouseup', schedulePosition);
        window.addEventListener('pagehide', savePosition);
        window.addEventListener('beforeunload', savePosition);
        warmNeighbors();
      }
    },
    restoreSession(focus = false, skip = false) {
      if (!model.enabled) return;
      const bookmark = initialBookmark; initialBookmark = null;
      restoring = true;
      // Run after Vditor's own page-local focus restoration; the document anchor takes precedence.
      requestAnimationFrame(() => requestAnimationFrame(() => {
        if (bookmark && !skip) restorePosition(bookmark, focus);
        restoring = false;
        bar.querySelectorAll('button, input, select').forEach(control => { control.disabled = false; });
        bar.setAttribute('aria-busy', 'false'); refresh(); schedulePosition();
      }));
    },
    update(text) {
      if (model.text === text) return;
      const bookmark = capturePosition();
      model.reset(text); searchOffset = -1;
      render(model.indexAt(resolveAnchor(text, bookmark?.anchor))); restorePosition(bookmark);
    },
    // Used for internal heading links; resolve headings against the whole source before scrolling.
    goto(fragment) {
      if (!model.enabled || !fragment) return;
      const needle = decodeURIComponent(fragment.replace(/^#/, '')).toLocaleLowerCase();
      const lines = model.text.matchAll(/^#{1,6}\s+(.+)$/gm);
      for (const line of lines) {
        const title = line[1].trim().replace(/\s+#+$/, '');
        const slug = title.toLocaleLowerCase().replace(/[^\p{L}\p{N}_\s-]/gu, '').replace(/\s+/g, '-');
        if (slug === needle || title.toLocaleLowerCase() === needle) { go(model.indexAt(line.index)); return; }
      }
    },
  };
}
