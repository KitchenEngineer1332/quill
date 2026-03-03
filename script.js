/**
 * QUILL EDITOR — JavaScript Core
 * Production-ready rich text editor
 * Pure ES6+ Vanilla JS, modular architecture
 */

'use strict';

/* ═══════════════════════════════════════════════════════
   1. CONSTANTS & CONFIG
═══════════════════════════════════════════════════════ */
const CONFIG = {
  AUTOSAVE_KEY:  'quill_autosave',
  DOC_KEY:       'quill_doc',
  SETTINGS_KEY:  'quill_settings',
  RECENT_KEY:    'quill_recent',
  ZOOM_MIN:      0.5,
  ZOOM_MAX:      2.0,
  ZOOM_STEP:     0.1,
  AUTOSAVE_MS:   2000,
  MAX_RECENT:    8,
};

/* ═══════════════════════════════════════════════════════
   2. STATE
═══════════════════════════════════════════════════════ */
const State = {
  theme:        'light',
  zoom:         1.0,
  isSaved:      false,
  autosaveTimer: null,
  findMatches:  [],
  findIndex:    0,
  savedRange:   null, // for modal operations that need saved selection
};

/* ═══════════════════════════════════════════════════════
   3. DOM REFERENCES
═══════════════════════════════════════════════════════ */
const $ = id => document.getElementById(id);
const $$ = sel => document.querySelectorAll(sel);

const DOM = {
  editor:         $('editor'),
  docTitle:       $('doc-title'),
  // Toolbar selects (font picker is now custom — see FontPicker module)
  tbSize:         $('tb-size'),
  tbColor:        $('tb-color'),
  tbHighlight:    $('tb-highlight'),
  colorInd:       $('color-indicator'),
  highlightInd:   $('highlight-indicator'),
  tbLineSpacing:  $('tb-linespacing'),
  // Toolbar buttons
  tbUndo:         $('tb-undo'),
  tbRedo:         $('tb-redo'),
  tbSave:         $('tb-save'),
  tbPrint:        $('tb-print'),
  tbLink:         $('tb-link'),
  tbImage:        $('tb-image'),
  tbTable:        $('tb-table'),
  tbPageBreak:    $('tb-pagebreak'),
  // Status bar
  statusMsg:      $('status-msg'),
  statusWords:    $('status-words'),
  statusChars:    $('status-chars'),
  statusAutosave: $('status-autosave'),
  statusZoom:     $('status-zoom'),
  // Sidebar
  zoomIn:         $('zoom-in'),
  zoomOut:        $('zoom-out'),
  zoomValue:      $('zoom-value'),
  recentList:     $('recent-list'),
  metaAuthor:     $('meta-author'),
  metaDate:       $('meta-date'),
  // Canvas
  editorCanvas:   $('editor-canvas'),
  pageWrapper:    document.querySelector('.page-wrapper'),
  dropIndicator:  $('drop-indicator'),
  // Float toolbar
  floatToolbar:   $('float-toolbar'),
  // Modals
  backdrop:       $('modal-backdrop'),
  // Find modal
  findInput:      $('find-input'),
  replaceInput:   $('replace-input'),
  findInfo:       $('find-info'),
  // Link modal
  linkText:       $('link-text'),
  linkUrl:        $('link-url'),
  // Image modal
  imageUrl:       $('image-url'),
  imageFile:      $('image-file'),
  // Table modal
  tableRows:      $('table-rows'),
  tableCols:      $('table-cols'),
  tableHeader:    $('table-header'),
  // File
  openFileInput:  $('open-file-input'),
};

/* ═══════════════════════════════════════════════════════
   4. UTILITY FUNCTIONS
═══════════════════════════════════════════════════════ */

/**
 * Save and restore selection across modal operations
 */
const Selection = {
  save() {
    const sel = window.getSelection();
    if (sel.rangeCount > 0) {
      State.savedRange = sel.getRangeAt(0).cloneRange();
    }
  },
  restore() {
    if (State.savedRange) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(State.savedRange);
    }
  },
  get text() {
    return window.getSelection()?.toString() || '';
  },
};

/**
 * Debounce utility
 */
function debounce(fn, delay) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

/**
 * Deep escape HTML for safe injection
 */
function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Count words in text
 */
function countWords(text) {
  const trimmed = text.trim();
  return trimmed ? trimmed.split(/\s+/).length : 0;
}

/**
 * Format date for display
 */
function formatDate(date) {
  return new Intl.DateTimeFormat('en-US', {
    year: 'numeric', month: 'short', day: 'numeric',
  }).format(date);
}

/* ═══════════════════════════════════════════════════════
   5. STATUS BAR
═══════════════════════════════════════════════════════ */
const StatusBar = {
  /**
   * Show a status message, revert after delay
   */
  show(msg, type = '', duration = 2500) {
    DOM.statusMsg.textContent = msg;
    DOM.statusMsg.className = `statusbar__msg${type ? ` is-${type}` : ''}`;
    if (duration > 0) {
      setTimeout(() => {
        DOM.statusMsg.textContent = 'Ready';
        DOM.statusMsg.className = 'statusbar__msg';
      }, duration);
    }
  },

  /**
   * Update word & character counts
   */
  updateCounts() {
    const text = DOM.editor.innerText || '';
    const words = countWords(text);
    const chars = text.replace(/\n/g, '').length;
    DOM.statusWords.textContent = `${words} ${words === 1 ? 'word' : 'words'}`;
    DOM.statusChars.textContent = `${chars} ${chars === 1 ? 'character' : 'characters'}`;
  },

  /**
   * Flash autosave indicator
   */
  flashAutosave() {
    DOM.statusAutosave.textContent = '✓ Autosaved';
    DOM.statusAutosave.classList.add('autosave-flash');
    setTimeout(() => {
      DOM.statusAutosave.classList.remove('autosave-flash');
      DOM.statusAutosave.textContent = '';
    }, 2200);
  },
};

/* ═══════════════════════════════════════════════════════
   6. THEME SYSTEM
═══════════════════════════════════════════════════════ */
const ThemeManager = {
  themes: ['light', 'dark', 'sepia'],

  apply(theme) {
    // Resolve 'system' to actual OS preference
    if (theme === 'system') {
      theme = window.matchMedia('(prefers-color-scheme: dark)').matches
        ? 'dark' : 'light';
    }

    // Validate
    if (!this.themes.includes(theme)) theme = 'light';

    State.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);

    // Update all theme buttons
    $$('.theme-btn').forEach(btn => {
      const isActive = btn.dataset.theme === theme;
      btn.classList.toggle('active', isActive);
      btn.setAttribute('aria-pressed', String(isActive));
    });

    // Persist
    this.save(theme);
    StatusBar.show(`Switched to ${theme} theme`, 'success');
  },

  save(theme) {
    try {
      const settings = this.loadSettings();
      settings.theme = theme;
      localStorage.setItem(CONFIG.SETTINGS_KEY, JSON.stringify(settings));
    } catch (e) { /* localStorage unavailable */ }
  },

  loadSettings() {
    try {
      return JSON.parse(localStorage.getItem(CONFIG.SETTINGS_KEY) || '{}');
    } catch { return {}; }
  },

  init() {
    // Listen for OS theme changes when using system theme
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
      const settings = this.loadSettings();
      if (settings.rawTheme === 'system') this.apply('system');
    });

    const settings = this.loadSettings();
    const saved = settings.theme || 'light';
    this.apply(saved);
  },
};

/* ═══════════════════════════════════════════════════════
   7. ZOOM MANAGER
═══════════════════════════════════════════════════════ */
const ZoomManager = {
  set(value) {
    State.zoom = Math.max(CONFIG.ZOOM_MIN, Math.min(CONFIG.ZOOM_MAX, value));
    State.zoom = Math.round(State.zoom * 10) / 10; // one decimal

    DOM.pageWrapper.style.setProperty('--zoom', State.zoom);
    // Compensate for CSS transform so scrolling still works
    const extra = (State.zoom - 1) * 1123;
    DOM.pageWrapper.style.marginBottom = `${Math.max(0, extra)}px`;

    const pct = `${Math.round(State.zoom * 100)}%`;
    DOM.zoomValue.textContent = pct;
    DOM.statusZoom.textContent = pct;

    // Persist
    try {
      const settings = ThemeManager.loadSettings();
      settings.zoom = State.zoom;
      localStorage.setItem(CONFIG.SETTINGS_KEY, JSON.stringify(settings));
    } catch (e) {}
  },

  in()    { this.set(State.zoom + CONFIG.ZOOM_STEP); },
  out()   { this.set(State.zoom - CONFIG.ZOOM_STEP); },
  reset() { this.set(1.0); },

  init() {
    const settings = ThemeManager.loadSettings();
    this.set(settings.zoom || 1.0);
  },
};

/* ═══════════════════════════════════════════════════════
   8. AUTOSAVE & RECENT DOCUMENTS
═══════════════════════════════════════════════════════ */
const Storage = {
  /**
   * Save editor content to localStorage (autosave)
   */
  autosave() {
    clearTimeout(State.autosaveTimer);
    State.autosaveTimer = setTimeout(() => {
      try {
        const content = DOM.editor.innerHTML;
        if (content && content !== '<p><br></p>') {
          localStorage.setItem(CONFIG.AUTOSAVE_KEY, content);
          StatusBar.flashAutosave();
        }
      } catch (e) { /* storage full or unavailable */ }
    }, CONFIG.AUTOSAVE_MS);
  },

  /**
   * Save named document
   */
  saveDocument(name, html) {
    try {
      const docs = this.getDocs();
      docs[name] = { html, savedAt: new Date().toISOString() };
      localStorage.setItem(CONFIG.DOC_KEY, JSON.stringify(docs));
      this.addRecent(name);
      return true;
    } catch (e) {
      return false;
    }
  },

  getDocs() {
    try { return JSON.parse(localStorage.getItem(CONFIG.DOC_KEY) || '{}'); }
    catch { return {}; }
  },

  /**
   * Add to recent documents list
   */
  addRecent(name) {
    try {
      let recent = JSON.parse(localStorage.getItem(CONFIG.RECENT_KEY) || '[]');
      recent = [name, ...recent.filter(r => r !== name)].slice(0, CONFIG.MAX_RECENT);
      localStorage.setItem(CONFIG.RECENT_KEY, JSON.stringify(recent));
      UI.renderRecent();
    } catch (e) {}
  },

  getRecent() {
    try { return JSON.parse(localStorage.getItem(CONFIG.RECENT_KEY) || '[]'); }
    catch { return []; }
  },

  /**
   * Load autosave on startup
   */
  loadAutosave() {
    try {
      const saved = localStorage.getItem(CONFIG.AUTOSAVE_KEY);
      if (saved && saved.trim()) {
        DOM.editor.innerHTML = saved;
        Editor.ensureContent();
        StatusBar.show('Autosave restored', 'success');
      }
    } catch (e) {}
  },
};

/* ═══════════════════════════════════════════════════════
   9. EDITOR CORE
═══════════════════════════════════════════════════════ */
const Editor = {
  /**
   * Execute a document command
   */
  exec(cmd, value = null) {
    DOM.editor.focus();
    try {
      document.execCommand(cmd, false, value);
    } catch (e) {
      console.warn(`execCommand failed: ${cmd}`, e);
    }
    this.updateToolbarState();
  },

  /**
   * Ensure editor always has at least one paragraph
   */
  ensureContent() {
    if (!DOM.editor.innerHTML.trim()) {
      DOM.editor.innerHTML = '<p><br></p>';
    }
  },

  /**
   * Update toolbar button active states based on current selection
   */
  updateToolbarState() {
    const commands = ['bold', 'italic', 'underline', 'strikeThrough',
      'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull',
      'insertUnorderedList', 'insertOrderedList'];

    commands.forEach(cmd => {
      const btn = document.querySelector(`[data-exec="${cmd}"]`);
      if (!btn) return;
      try {
        btn.classList.toggle('is-active', document.queryCommandState(cmd));
      } catch (e) {}
    });

    // Sync font picker to cursor position
    FontPicker.syncFromCursor();
  },

  /**
   * Apply line spacing to current selection or cursor paragraph
   */
  applyLineSpacing(value) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;

    const range = sel.getRangeAt(0);
    let container = range.commonAncestorContainer;

    // Walk up to find block element
    while (container && container !== DOM.editor &&
      !['P','DIV','LI','H1','H2','H3','H4','H5','H6','BLOCKQUOTE'].includes(container.nodeName)) {
      container = container.parentNode;
    }

    if (container && container !== DOM.editor) {
      container.style.lineHeight = value;
    } else {
      // Apply to all paragraphs in selection
      DOM.editor.querySelectorAll('p, div, li').forEach(el => {
        el.style.lineHeight = value;
      });
    }
  },

  /**
   * Clear all formatting from selection
   */
  clearFormat() {
    this.exec('removeFormat');
    this.exec('unlink');
  },

  /**
   * Get clean text content
   */
  getText() {
    return DOM.editor.innerText || '';
  },

  /**
   * Get HTML content
   */
  getHtml() {
    return DOM.editor.innerHTML || '';
  },
};

/* ═══════════════════════════════════════════════════════
   10. FIND & REPLACE
═══════════════════════════════════════════════════════ */
const FindReplace = {
  _originalContent: null,

  /**
   * Clear all highlights
   */
  clearHighlights() {
    if (!this._originalContent) return;
    DOM.editor.innerHTML = this._originalContent;
    this._originalContent = null;
    State.findMatches = [];
    State.findIndex = 0;
    DOM.findInfo.textContent = '';
  },

  /**
   * Highlight all occurrences of query
   */
  highlight(query) {
    this.clearHighlights();
    if (!query) return;

    this._originalContent = DOM.editor.innerHTML;
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, 'gi');

    // Walk text nodes and wrap matches
    const walker = document.createTreeWalker(
      DOM.editor,
      NodeFilter.SHOW_TEXT,
      null
    );

    const textNodes = [];
    let node;
    while ((node = walker.nextNode())) textNodes.push(node);

    textNodes.forEach(textNode => {
      if (!textNode.nodeValue.trim()) return;
      const parent = textNode.parentNode;
      if (!parent || parent.closest('.find-highlight')) return;

      const parts = textNode.nodeValue.split(regex);
      if (parts.length <= 1) return;

      const matches = textNode.nodeValue.match(regex);
      const frag = document.createDocumentFragment();

      parts.forEach((part, i) => {
        if (part) frag.appendChild(document.createTextNode(part));
        if (matches && matches[i]) {
          const mark = document.createElement('mark');
          mark.className = 'find-highlight';
          mark.textContent = matches[i];
          frag.appendChild(mark);
        }
      });

      parent.replaceChild(frag, textNode);
    });

    State.findMatches = Array.from(DOM.editor.querySelectorAll('.find-highlight'));
    State.findIndex = 0;
    this.scrollToMatch(0);
    DOM.findInfo.textContent = `${State.findMatches.length} match${State.findMatches.length !== 1 ? 'es' : ''} found`;
  },

  /**
   * Scroll to and highlight current match
   */
  scrollToMatch(idx) {
    State.findMatches.forEach(el => el.classList.remove('current'));
    if (State.findMatches.length === 0) return;

    const el = State.findMatches[idx];
    if (!el) return;
    el.classList.add('current');
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  },

  next() {
    if (!State.findMatches.length) return;
    State.findIndex = (State.findIndex + 1) % State.findMatches.length;
    this.scrollToMatch(State.findIndex);
  },

  prev() {
    if (!State.findMatches.length) return;
    State.findIndex = (State.findIndex - 1 + State.findMatches.length) % State.findMatches.length;
    this.scrollToMatch(State.findIndex);
  },

  /**
   * Replace current match
   */
  replaceCurrent(replacement) {
    if (!State.findMatches.length) return;
    const el = State.findMatches[State.findIndex];
    if (!el) return;
    el.replaceWith(document.createTextNode(replacement));
    State.findMatches.splice(State.findIndex, 1);
    if (State.findMatches.length) {
      State.findIndex = State.findIndex % State.findMatches.length;
      this.scrollToMatch(State.findIndex);
    }
    DOM.findInfo.textContent = `${State.findMatches.length} match${State.findMatches.length !== 1 ? 'es' : ''} remaining`;
  },

  /**
   * Replace all matches
   */
  replaceAll(query, replacement) {
    this.clearHighlights();
    if (!query) return;

    DOM.editor.innerHTML = DOM.editor.innerHTML.replace(
      new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'),
      replacement
    );

    const count = DOM.editor.innerHTML
      .split(new RegExp(replacement.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i')).length - 1;

    StatusBar.show(`Replaced all occurrences`, 'success');
    DOM.findInfo.textContent = `All instances replaced`;
    Storage.autosave();
  },
};

/* ═══════════════════════════════════════════════════════
   11. FILE OPERATIONS
═══════════════════════════════════════════════════════ */
const FileOps = {
  /**
   * Create a new document
   */
  newDocument() {
    if (DOM.editor.innerHTML && DOM.editor.innerHTML !== '<p><br></p>') {
      if (!confirm('Start a new document? Unsaved changes will be lost.')) return;
    }
    DOM.editor.innerHTML = '<p><br></p>';
    DOM.docTitle.value = 'Untitled Document';
    DOM.metaDate.value = formatDate(new Date());
    StatusBar.show('New document created', 'success');
    Storage.autosave();
    DOM.editor.focus();
  },

  /**
   * Save as HTML file download
   */
  saveAsHtml() {
    const title = DOM.docTitle.value || 'document';
    const author = DOM.metaAuthor.value || '';
    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>${escapeHtml(title)}</title>
  <meta name="author" content="${escapeHtml(author)}">
  <style>
    body { font-family: Georgia, serif; max-width: 800px; margin: 60px auto; padding: 0 40px; line-height: 1.7; color: #1a1a2e; }
    table { border-collapse: collapse; width: 100%; }
    td, th { border: 1px solid #ccc; padding: 8px 12px; }
    th { background: #f5f5f0; }
    img { max-width: 100%; }
    .page-break { page-break-after: always; }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  ${DOM.editor.innerHTML}
</body>
</html>`;

    this._download(`${title}.html`, html, 'text/html');
    StatusBar.show('Saved as HTML', 'success');
    Storage.saveDocument(title, DOM.editor.innerHTML);
  },

  /**
   * Save as plain text
   */
  saveAsTxt() {
    const title = DOM.docTitle.value || 'document';
    const text = DOM.editor.innerText || '';
    this._download(`${title}.txt`, text, 'text/plain');
    StatusBar.show('Saved as TXT', 'success');
  },

  /**
   * Export as PDF via print dialog
   */
  exportPdf() {
    StatusBar.show('Preparing PDF export…');
    setTimeout(() => window.print(), 100);
  },

  /**
   * Open file from disk
   */
  openFile(file) {
    if (!file) return;
    const reader = new FileReader();
    const ext = file.name.split('.').pop().toLowerCase();

    reader.onload = (e) => {
      const content = e.target.result;
      if (ext === 'html' || ext === 'htm') {
        // Extract body content for safety
        const tmp = document.createElement('div');
        tmp.innerHTML = content;
        const body = tmp.querySelector('body');
        DOM.editor.innerHTML = body ? body.innerHTML : content;
      } else {
        // Plain text — wrap in paragraphs
        DOM.editor.innerHTML = content
          .split('\n')
          .map(line => `<p>${escapeHtml(line) || '<br>'}</p>`)
          .join('');
      }
      DOM.docTitle.value = file.name.replace(/\.[^.]+$/, '');
      Editor.ensureContent();
      StatusBar.show(`Opened "${file.name}"`, 'success');
      StatusBar.updateCounts();
      Storage.addRecent(file.name);
    };

    reader.onerror = () => StatusBar.show('Failed to open file', 'error');
    reader.readAsText(file);
  },

  /**
   * Trigger helper: create and click a download link
   */
  _download(filename, content, type) {
    const blob = new Blob([content], { type: `${type};charset=utf-8` });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  },
};

/* ═══════════════════════════════════════════════════════
   12. MODAL SYSTEM
═══════════════════════════════════════════════════════ */
const Modal = {
  _activeModal: null,

  open(id) {
    // Close any open modal first
    if (this._activeModal) this.close();

    const el = $(`modal-${id}`);
    if (!el) return;

    el.removeAttribute('hidden');
    el.setAttribute('aria-hidden', 'false');
    DOM.backdrop.classList.add('is-open');
    DOM.backdrop.setAttribute('aria-hidden', 'false');

    // Animate in
    requestAnimationFrame(() => {
      el.classList.add('is-open');
    });

    this._activeModal = el;

    // Focus first input in modal
    setTimeout(() => {
      const firstInput = el.querySelector('input:not([type="checkbox"]):not([type="file"]), textarea, button.modal__btn--primary');
      if (firstInput) firstInput.focus();
    }, 80);
  },

  close() {
    if (!this._activeModal) return;
    const el = this._activeModal;

    el.classList.remove('is-open');
    DOM.backdrop.classList.remove('is-open');

    setTimeout(() => {
      el.setAttribute('hidden', '');
      el.setAttribute('aria-hidden', 'true');
      DOM.backdrop.setAttribute('aria-hidden', 'true');
    }, 200);

    this._activeModal = null;
    DOM.editor.focus();
  },
};

/* ═══════════════════════════════════════════════════════
   13. INSERT OPERATIONS
═══════════════════════════════════════════════════════ */
const Insert = {
  /**
   * Insert a hyperlink
   */
  link(text, url) {
    if (!url) return;
    const displayText = text || url;
    const html = `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(displayText)}</a>`;

    Selection.restore();
    if (Selection.text) {
      Editor.exec('createLink', url);
    } else {
      Editor.exec('insertHTML', html);
    }
  },

  /**
   * Insert an image from URL
   */
  imageFromUrl(url) {
    if (!url) return;
    const html = `<img src="${escapeHtml(url)}" alt="Image" style="max-width:100%">`;
    Editor.exec('insertHTML', html);
  },

  /**
   * Insert an image from file (base64)
   */
  imageFromFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      Selection.restore();
      this.imageFromUrl(e.target.result);
    };
    reader.readAsDataURL(file);
  },

  /**
   * Insert a table
   */
  table(rows, cols, hasHeader) {
    let html = '<table border="1" style="border-collapse:collapse;width:100%;margin:1em 0">';

    if (hasHeader) {
      html += '<thead><tr>';
      for (let c = 0; c < cols; c++) {
        html += `<th contenteditable="true" style="padding:8px 12px;background:#f5f5f0;font-weight:600">Header ${c + 1}</th>`;
      }
      html += '</tr></thead>';
    }

    html += '<tbody>';
    const dataRows = hasHeader ? rows - 1 : rows;
    for (let r = 0; r < Math.max(1, dataRows); r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) {
        html += `<td contenteditable="true" style="padding:8px 12px">&nbsp;</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';

    Editor.exec('insertHTML', html);
  },

  /**
   * Insert page break
   */
  pageBreak() {
    Editor.exec('insertHTML', '<div class="page-break"></div><p><br></p>');
  },
};

/* ═══════════════════════════════════════════════════════
   14. FLOATING TOOLBAR
═══════════════════════════════════════════════════════ */
const FloatToolbar = {
  showTimeout: null,

  /**
   * Show or hide the floating toolbar based on selection
   */
  update() {
    clearTimeout(this.showTimeout);
    const sel = window.getSelection();

    if (!sel || sel.isCollapsed || !sel.toString().trim() || !DOM.editor.contains(sel.anchorNode)) {
      this.hide();
      return;
    }

    this.showTimeout = setTimeout(() => this.show(sel), 80);
  },

  show(sel) {
    const range = sel.getRangeAt(0);
    const rect = range.getBoundingClientRect();
    if (!rect.width) return;

    const tb = DOM.floatToolbar;
    tb.removeAttribute('hidden');
    tb.setAttribute('aria-hidden', 'false');

    // Position above selection
    const tbW = 220;
    let left = rect.left + rect.width / 2 - tbW / 2;
    let top  = rect.top - 48;

    // Clamp to viewport
    left = Math.max(8, Math.min(left, window.innerWidth - tbW - 8));
    top  = top < 8 ? rect.bottom + 8 : top;

    tb.style.left = `${left}px`;
    tb.style.top  = `${top}px`;

    requestAnimationFrame(() => tb.classList.add('is-visible'));
  },

  hide() {
    DOM.floatToolbar.classList.remove('is-visible');
    setTimeout(() => {
      DOM.floatToolbar.setAttribute('hidden', '');
      DOM.floatToolbar.setAttribute('aria-hidden', 'true');
    }, 180);
  },
};

/* ═══════════════════════════════════════════════════════
   15. MENU BAR
═══════════════════════════════════════════════════════ */
const MenuBar = {
  _openMenu: null,

  init() {
    const triggers = $$('.menu__trigger');

    triggers.forEach(trigger => {
      trigger.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = trigger.dataset.menu;
        const dropdown = $(`menu-${key}`);
        const isOpen = dropdown.classList.contains('is-open');

        this.closeAll();
        if (!isOpen) this.open(trigger, dropdown);
      });
    });

    // Close on outside click
    document.addEventListener('click', () => this.closeAll());

    // Keyboard nav for menu items
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') this.closeAll();
    });
  },

  open(trigger, dropdown) {
    trigger.classList.add('is-open');
    trigger.setAttribute('aria-expanded', 'true');
    dropdown.classList.add('is-open');
    this._openMenu = { trigger, dropdown };
  },

  closeAll() {
    $$('.menu__trigger.is-open').forEach(t => {
      t.classList.remove('is-open');
      t.setAttribute('aria-expanded', 'false');
    });
    $$('.menu__dropdown.is-open').forEach(d => d.classList.remove('is-open'));
    this._openMenu = null;
  },
};

/* ═══════════════════════════════════════════════════════
   16. UI ORCHESTRATION
═══════════════════════════════════════════════════════ */
const UI = {
  /**
   * Render recent documents sidebar list
   */
  renderRecent() {
    const recent = Storage.getRecent();
    if (!recent.length) {
      DOM.recentList.innerHTML = '<li class="sidebar__recent-empty">No recent documents</li>';
      return;
    }

    DOM.recentList.innerHTML = recent.map(name =>
      `<li class="sidebar__recent-item" tabindex="0" role="button" aria-label="Open ${escapeHtml(name)}">${escapeHtml(name)}</li>`
    ).join('');

    // Clicking a recent item loads it
    DOM.recentList.querySelectorAll('.sidebar__recent-item').forEach((item, i) => {
      const action = () => {
        const docs = Storage.getDocs();
        const doc = docs[recent[i]];
        if (doc) {
          DOM.editor.innerHTML = doc.html;
          DOM.docTitle.value = recent[i];
          Editor.ensureContent();
          StatusBar.show(`Loaded "${recent[i]}"`, 'success');
          StatusBar.updateCounts();
        }
      };
      item.addEventListener('click', action);
      item.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') action(); });
    });
  },
};

/* ═══════════════════════════════════════════════════════
   17. KEYBOARD SHORTCUTS
═══════════════════════════════════════════════════════ */
function initKeyboardShortcuts() {
  document.addEventListener('keydown', (e) => {
    const ctrl = e.ctrlKey || e.metaKey;

    if (ctrl) {
      switch (e.key.toLowerCase()) {
        case 'b': e.preventDefault(); Editor.exec('bold'); break;
        case 'i': e.preventDefault(); Editor.exec('italic'); break;
        case 'u': e.preventDefault(); Editor.exec('underline'); break;
        case 'z':
          e.preventDefault();
          if (e.shiftKey) Editor.exec('redo');
          else Editor.exec('undo');
          break;
        case 'y': e.preventDefault(); Editor.exec('redo'); break;
        case 'a': e.preventDefault(); Editor.exec('selectAll'); break;
        case 's':
          e.preventDefault();
          FileOps.saveAsHtml();
          break;
        case 'n':
          e.preventDefault();
          FileOps.newDocument();
          break;
        case 'o':
          e.preventDefault();
          DOM.openFileInput.click();
          break;
        case 'p':
          e.preventDefault();
          FileOps.exportPdf();
          break;
        case 'f':
          e.preventDefault();
          Modal.open('find');
          break;
        case '=':
        case '+':
          e.preventDefault();
          ZoomManager.in();
          break;
        case '-':
          e.preventDefault();
          ZoomManager.out();
          break;
        case '0':
          e.preventDefault();
          ZoomManager.reset();
          break;
        case '\\':
          e.preventDefault();
          Editor.clearFormat();
          break;
      }

      // Strikethrough: Ctrl+Shift+X
      if (e.shiftKey && e.key === 'X') {
        e.preventDefault();
        Editor.exec('strikeThrough');
      }
    }

    if (e.key === 'F11') {
      e.preventDefault();
      toggleFullscreen();
    }

    // Tab in editor → insert 4 spaces (prevent focus trap)
    if (e.key === 'Tab' && document.activeElement === DOM.editor) {
      e.preventDefault();
      Editor.exec('insertHTML', '&nbsp;&nbsp;&nbsp;&nbsp;');
    }
  });
}

/* ═══════════════════════════════════════════════════════
   18. DRAG & DROP
═══════════════════════════════════════════════════════ */
function initDragDrop() {
  const canvas = DOM.editorCanvas;

  ['dragenter', 'dragover'].forEach(evt =>
    canvas.addEventListener(evt, e => {
      e.preventDefault();
      e.stopPropagation();
      DOM.dropIndicator.classList.add('is-visible');
    })
  );

  ['dragleave', 'dragend'].forEach(evt =>
    canvas.addEventListener(evt, (e) => {
      if (!canvas.contains(e.relatedTarget)) {
        DOM.dropIndicator.classList.remove('is-visible');
      }
    })
  );

  canvas.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    DOM.dropIndicator.classList.remove('is-visible');

    const files = Array.from(e.dataTransfer?.files || []);
    const textFile = files.find(f => /text\/(plain|html)/.test(f.type) || /\.(txt|html|htm)$/i.test(f.name));
    const imageFile = files.find(f => f.type.startsWith('image/'));

    if (textFile) {
      FileOps.openFile(textFile);
    } else if (imageFile) {
      Insert.imageFromFile(imageFile);
    } else if (files.length) {
      StatusBar.show('Unsupported file type', 'error');
    }

    // Text drop
    const text = e.dataTransfer?.getData('text/plain');
    if (!files.length && text) {
      Editor.exec('insertText', text);
    }
  });
}

/* ═══════════════════════════════════════════════════════
   19. FULLSCREEN
═══════════════════════════════════════════════════════ */
function toggleFullscreen() {
  if (!document.fullscreenElement) {
    document.documentElement.requestFullscreen?.();
    StatusBar.show('Fullscreen — press F11 or Esc to exit');
  } else {
    document.exitFullscreen?.();
  }
}

/* ═══════════════════════════════════════════════════════
   20. EVENT WIRING — TOOLBAR
═══════════════════════════════════════════════════════ */
function initToolbar() {
  // Exec buttons (data-exec attribute)
  $$('[data-exec]').forEach(btn => {
    btn.addEventListener('mousedown', e => e.preventDefault()); // keep selection
    btn.addEventListener('click', () => {
      Editor.exec(btn.dataset.exec);
    });
  });

  // Font size
  DOM.tbSize.addEventListener('change', () => {
    Editor.exec('fontSize', DOM.tbSize.value);
  });

  // Text color
  DOM.tbColor.addEventListener('input', () => {
    const c = DOM.tbColor.value;
    DOM.colorInd.style.background = c;
    Editor.exec('foreColor', c);
  });

  // Highlight color
  DOM.tbHighlight.addEventListener('input', () => {
    const c = DOM.tbHighlight.value;
    DOM.highlightInd.style.background = c;
    Editor.exec('hiliteColor', c);
  });

  // Line spacing
  DOM.tbLineSpacing.addEventListener('change', () => {
    Editor.applyLineSpacing(DOM.tbLineSpacing.value);
  });

  // Undo / Redo
  DOM.tbUndo.addEventListener('click', () => Editor.exec('undo'));
  DOM.tbRedo.addEventListener('click', () => Editor.exec('redo'));

  // Save
  DOM.tbSave.addEventListener('click', () => FileOps.saveAsHtml());

  // Print
  DOM.tbPrint.addEventListener('click', () => FileOps.exportPdf());

  // Link
  DOM.tbLink.addEventListener('click', () => {
    Selection.save();
    DOM.linkText.value = Selection.text;
    DOM.linkUrl.value  = '';
    Modal.open('link');
  });

  // Image
  DOM.tbImage.addEventListener('click', () => {
    Selection.save();
    DOM.imageUrl.value = '';
    DOM.imageFile.value = '';
    Modal.open('image');
  });

  // Table
  DOM.tbTable.addEventListener('click', () => {
    Selection.save();
    Modal.open('table');
  });

  // Page break
  DOM.tbPageBreak.addEventListener('click', () => Insert.pageBreak());

  // Zoom
  DOM.zoomIn.addEventListener('click',  () => ZoomManager.in());
  DOM.zoomOut.addEventListener('click', () => ZoomManager.out());
  $('status-zoom-in').addEventListener('click',  () => ZoomManager.in());
  $('status-zoom-out').addEventListener('click', () => ZoomManager.out());
}

/* ═══════════════════════════════════════════════════════
   21. EVENT WIRING — MENU BAR
═══════════════════════════════════════════════════════ */
function initMenuBar() {
  // File
  $('cmd-new')?.addEventListener('click',   () => { MenuBar.closeAll(); FileOps.newDocument(); });
  $('cmd-open')?.addEventListener('click',  () => { MenuBar.closeAll(); DOM.openFileInput.click(); });
  $('cmd-save-html')?.addEventListener('click', () => { MenuBar.closeAll(); FileOps.saveAsHtml(); });
  $('cmd-save-txt')?.addEventListener('click',  () => { MenuBar.closeAll(); FileOps.saveAsTxt(); });
  $('cmd-export-pdf')?.addEventListener('click',() => { MenuBar.closeAll(); FileOps.exportPdf(); });
  $('cmd-recent')?.addEventListener('click',    () => { MenuBar.closeAll(); StatusBar.show('Recent docs shown in sidebar'); });

  // Edit
  $('cmd-undo')?.addEventListener('click',       () => { MenuBar.closeAll(); Editor.exec('undo'); });
  $('cmd-redo')?.addEventListener('click',       () => { MenuBar.closeAll(); Editor.exec('redo'); });
  $('cmd-cut')?.addEventListener('click',        () => { MenuBar.closeAll(); Editor.exec('cut'); });
  $('cmd-copy')?.addEventListener('click',       () => { MenuBar.closeAll(); Editor.exec('copy'); });
  $('cmd-paste')?.addEventListener('click',      () => { MenuBar.closeAll(); Editor.exec('paste'); });
  $('cmd-select-all')?.addEventListener('click', () => { MenuBar.closeAll(); Editor.exec('selectAll'); });
  $('cmd-find')?.addEventListener('click',       () => { MenuBar.closeAll(); Modal.open('find'); });

  // View
  $('cmd-zoom-in')?.addEventListener('click',    () => { MenuBar.closeAll(); ZoomManager.in(); });
  $('cmd-zoom-out')?.addEventListener('click',   () => { MenuBar.closeAll(); ZoomManager.out(); });
  $('cmd-zoom-reset')?.addEventListener('click', () => { MenuBar.closeAll(); ZoomManager.reset(); });
  $('cmd-fullscreen')?.addEventListener('click', () => { MenuBar.closeAll(); toggleFullscreen(); });

  // Theme from menu
  $$('[data-theme-btn]').forEach(btn => {
    btn.addEventListener('click', () => {
      MenuBar.closeAll();
      ThemeManager.apply(btn.dataset.themeBtn);
    });
  });

  // Format
  $('cmd-insert-link')?.addEventListener('click', () => {
    MenuBar.closeAll();
    Selection.save();
    DOM.linkText.value = Selection.text;
    DOM.linkUrl.value  = '';
    Modal.open('link');
  });
  $('cmd-insert-image')?.addEventListener('click', () => {
    MenuBar.closeAll();
    Modal.open('image');
  });
  $('cmd-insert-table')?.addEventListener('click', () => {
    MenuBar.closeAll();
    Modal.open('table');
  });
  $('cmd-clear-format')?.addEventListener('click', () => {
    MenuBar.closeAll();
    Editor.clearFormat();
  });

  // Help
  $('cmd-shortcuts')?.addEventListener('click', () => { MenuBar.closeAll(); Modal.open('shortcuts'); });
  $('cmd-about')?.addEventListener('click',     () => { MenuBar.closeAll(); Modal.open('about'); });

  // Theme buttons in menubar header
  $$('.theme-btn').forEach(btn => {
    btn.addEventListener('click', () => ThemeManager.apply(btn.dataset.theme));
  });

  // File open
  DOM.openFileInput.addEventListener('change', (e) => {
    FileOps.openFile(e.target.files[0]);
    e.target.value = ''; // allow reopening same file
  });
}

/* ═══════════════════════════════════════════════════════
   22. EVENT WIRING — MODALS
═══════════════════════════════════════════════════════ */
function initModals() {
  // Generic close buttons
  $$('[data-close-modal]').forEach(btn => {
    btn.addEventListener('click', () => Modal.close());
  });

  // Backdrop close
  DOM.backdrop.addEventListener('click', () => Modal.close());

  // Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && Modal._activeModal) Modal.close();
  });

  // ── FIND & REPLACE ──
  // Highlight as-you-type (find-btn doesn't exist in HTML; input event is better UX anyway)
  DOM.findInput.addEventListener('input', () => {
    FindReplace.highlight(DOM.findInput.value);
  });

  DOM.findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (State.findMatches.length) FindReplace.next();
      else FindReplace.highlight(DOM.findInput.value);
    }
  });

  $('find-prev')?.addEventListener('click', () => FindReplace.prev());
  $('find-next')?.addEventListener('click', () => FindReplace.next());

  $('replace-one')?.addEventListener('click', () => {
    FindReplace.replaceCurrent(DOM.replaceInput.value);
    Storage.autosave();
  });

  $('replace-all')?.addEventListener('click', () => {
    FindReplace.replaceAll(DOM.findInput.value, DOM.replaceInput.value);
  });

  // Clear highlights when modal closes
  const origClose = Modal.close.bind(Modal);
  Modal.close = function () {
    // Clear find highlights when closing find modal
    if (Modal._activeModal?.id === 'modal-find') {
      FindReplace.clearHighlights();
    }
    origClose();
  };

  // ── LINK INSERT ──
  $('link-insert')?.addEventListener('click', () => {
    Insert.link(DOM.linkText.value, DOM.linkUrl.value);
    Modal.close();
    Storage.autosave();
  });

  DOM.linkUrl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') $('link-insert')?.click();
  });

  // ── IMAGE INSERT ──
  $('image-insert')?.addEventListener('click', () => {
    const url  = DOM.imageUrl.value.trim();
    const file = DOM.imageFile.files[0];

    if (file) {
      Insert.imageFromFile(file);
    } else if (url) {
      Selection.restore();
      Insert.imageFromUrl(url);
    } else {
      StatusBar.show('Please provide an image URL or select a file', 'error');
      return;
    }
    Modal.close();
    Storage.autosave();
  });

  // ── TABLE INSERT ──
  $('table-insert')?.addEventListener('click', () => {
    const rows = Math.max(1, Math.min(20, parseInt(DOM.tableRows.value) || 3));
    const cols = Math.max(1, Math.min(10, parseInt(DOM.tableCols.value) || 3));
    Selection.restore();
    Insert.table(rows, cols, DOM.tableHeader.checked);
    Modal.close();
    Storage.autosave();
  });

  // ── FLOAT TOOLBAR LINK ──
  $('float-link')?.addEventListener('click', () => {
    Selection.save();
    DOM.linkText.value = Selection.text;
    DOM.linkUrl.value  = '';
    FloatToolbar.hide();
    Modal.open('link');
  });
}

/* ═══════════════════════════════════════════════════════
   23. EVENT WIRING — EDITOR
═══════════════════════════════════════════════════════ */
function initEditor() {
  // Prevent default on toolbar mousedowns to keep selection
  $$('.toolbar__btn, .toolbar__select, .float-toolbar__btn').forEach(el => {
    el.addEventListener('mousedown', e => {
      if (el.tagName !== 'SELECT') e.preventDefault();
    });
  });

  // Content events
  DOM.editor.addEventListener('input', debounce(() => {
    Editor.ensureContent();
    StatusBar.updateCounts();
    Storage.autosave();
    Editor.updateToolbarState();
  }, 200));

  // Selection change for floating toolbar and toolbar state
  document.addEventListener('selectionchange', () => {
    FloatToolbar.update();
    Editor.updateToolbarState();
  });

  // Initial content
  Editor.ensureContent();
  DOM.editor.focus();
}

/* ═══════════════════════════════════════════════════════
   24. FONT PICKER — Custom dropdown with live hover preview
   ─────────────────────────────────────────────────────
   Safe live-preview strategy:
   1. On open: snapshot editor.innerHTML + serialize selection as
      character offsets (survives innerHTML restore).
   2. On each hover: restore innerHTML snapshot → restore selection
      from offsets → execCommand('fontName', value).
      The user sees their text change font instantly.
   3. On cancel/close: restore innerHTML snapshot (text unchanged).
   4. On commit: keep current DOM state (last preview = final result),
      clear snapshot, autosave.
   This avoids all DOM surgery (extractContents / insertNode) which
   was the source of the text-deletion bug.
═══════════════════════════════════════════════════════ */
const FontPicker = (() => {

  /* ── Font data ── */
  const FONT_GROUPS = [
    { label: 'Sans-Serif', fonts: [
      { name: 'Barlow',            value: 'Barlow, sans-serif' },
      { name: 'DM Sans',           value: "'DM Sans', sans-serif" },
      { name: 'Figtree',           value: 'Figtree, sans-serif' },
      { name: 'IBM Plex Sans',     value: "'IBM Plex Sans', sans-serif" },
      { name: 'Inter',             value: 'Inter, sans-serif' },
      { name: 'Josefin Sans',      value: "'Josefin Sans', sans-serif" },
      { name: 'Lato',              value: 'Lato, sans-serif' },
      { name: 'Montserrat',        value: 'Montserrat, sans-serif' },
      { name: 'Mulish',            value: 'Mulish, sans-serif' },
      { name: 'Nunito',            value: 'Nunito, sans-serif' },
      { name: 'Open Sans',         value: "'Open Sans', sans-serif" },
      { name: 'Oswald',            value: 'Oswald, sans-serif' },
      { name: 'Outfit',            value: 'Outfit, sans-serif' },
      { name: 'Plus Jakarta Sans', value: "'Plus Jakarta Sans', sans-serif" },
      { name: 'Poppins',           value: 'Poppins, sans-serif' },
      { name: 'PT Sans',           value: "'PT Sans', sans-serif" },
      { name: 'Quicksand',         value: 'Quicksand, sans-serif' },
      { name: 'Raleway',           value: 'Raleway, sans-serif' },
      { name: 'Roboto',            value: 'Roboto, sans-serif' },
      { name: 'Sora',              value: 'Sora, sans-serif' },
      { name: 'Source Sans 3',     value: "'Source Sans 3', sans-serif" },
      { name: 'Ubuntu',            value: 'Ubuntu, sans-serif' },
      { name: 'Work Sans',         value: "'Work Sans', sans-serif" },
    ]},
    { label: 'Serif', fonts: [
      { name: 'Bitter',             value: 'Bitter, serif' },
      { name: 'Cormorant Garamond', value: "'Cormorant Garamond', serif" },
      { name: 'Crimson Pro',        value: "'Crimson Pro', serif" },
      { name: 'DM Serif Display',   value: "'DM Serif Display', serif" },
      { name: 'Domine',             value: 'Domine, serif' },
      { name: 'EB Garamond',        value: "'EB Garamond', serif" },
      { name: 'Georgia',            value: 'Georgia, serif' },
      { name: 'Libre Baskerville',  value: "'Libre Baskerville', serif" },
      { name: 'Lora',               value: 'Lora, serif' },
      { name: 'Merriweather',       value: 'Merriweather, serif' },
      { name: 'Playfair Display',   value: "'Playfair Display', serif" },
      { name: 'PT Serif',           value: "'PT Serif', serif" },
    ]},
    { label: 'Display', fonts: [
      { name: 'Abril Fatface', value: "'Abril Fatface', cursive" },
      { name: 'Alfa Slab One', value: "'Alfa Slab One', serif" },
      { name: 'Bebas Neue',    value: "'Bebas Neue', sans-serif" },
      { name: 'Fredoka',       value: 'Fredoka, sans-serif' },
      { name: 'Righteous',     value: 'Righteous, sans-serif' },
    ]},
    { label: 'Handwriting', fonts: [
      { name: 'Caveat',         value: 'Caveat, cursive' },
      { name: 'Dancing Script', value: "'Dancing Script', cursive" },
      { name: 'Great Vibes',    value: "'Great Vibes', cursive" },
      { name: 'Kalam',          value: 'Kalam, cursive' },
      { name: 'Lobster',        value: 'Lobster, cursive' },
      { name: 'Pacifico',       value: 'Pacifico, cursive' },
      { name: 'Patrick Hand',   value: "'Patrick Hand', cursive" },
      { name: 'Satisfy',        value: 'Satisfy, cursive' },
    ]},
    { label: 'Monospace', fonts: [
      { name: 'Fira Code',       value: "'Fira Code', monospace" },
      { name: 'IBM Plex Mono',   value: "'IBM Plex Mono', monospace" },
      { name: 'Inconsolata',     value: 'Inconsolata, monospace' },
      { name: 'JetBrains Mono',  value: "'JetBrains Mono', monospace" },
      { name: 'PT Mono',         value: "'PT Mono', monospace" },
      { name: 'Roboto Mono',     value: "'Roboto Mono', monospace" },
      { name: 'Source Code Pro', value: "'Source Code Pro', monospace" },
      { name: 'Space Mono',      value: "'Space Mono', monospace" },
      { name: 'Courier New',     value: "'Courier New', monospace" },
    ]},
    { label: 'System', fonts: [
      { name: 'Arial',           value: 'Arial, sans-serif' },
      { name: 'Times New Roman', value: "'Times New Roman', serif" },
      { name: 'Verdana',         value: 'Verdana, sans-serif' },
      { name: 'Trebuchet MS',    value: "'Trebuchet MS', sans-serif" },
      { name: 'Impact',          value: 'Impact, sans-serif' },
      { name: 'Comic Sans MS',   value: "'Comic Sans MS', cursive" },
    ]},
  ];

  /* ── Private state ── */
  let currentFont   = 'Georgia, serif';
  let currentName   = 'Georgia';
  let isOpen        = false;
  let allItems      = [];
  let _origHTML     = null;   // innerHTML snapshot taken when picker opens
  let _selOffsets   = null;   // {start, end} char offsets of selection at open
  let _hasPreviewed = false;  // true if at least one preview execCommand fired

  /* ── DOM refs (populated in init) ── */
  let triggerEl, labelEl, panelEl, listEl, searchEl;

  /* ─────────────────────────────────
     Selection serialization helpers
     Works on raw text character offsets so it survives innerHTML restores.
  ───────────────────────────────── */
  function _getSelectionOffsets() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    const range = sel.getRangeAt(0);
    if (!DOM.editor.contains(range.commonAncestorContainer)) return null;

    // Count chars from editor start → range start
    const preRange = document.createRange();
    preRange.selectNodeContents(DOM.editor);
    preRange.setEnd(range.startContainer, range.startOffset);
    const start = preRange.toString().length;
    const end   = start + range.toString().length;
    return { start, end };
  }

  function _restoreSelectionFromOffsets(offsets) {
    if (!offsets) return;
    const { start, end } = offsets;
    const walker = document.createTreeWalker(DOM.editor, NodeFilter.SHOW_TEXT, null);
    let charCount = 0, startNode = null, endNode = null, startOff = 0, endOff = 0;
    let node;
    while ((node = walker.nextNode())) {
      const len = node.length;
      if (!startNode && charCount + len >= start) {
        startNode = node;
        startOff  = start - charCount;
      }
      if (!endNode && charCount + len >= end) {
        endNode = node;
        endOff  = end - charCount;
        break;
      }
      charCount += len;
    }
    if (!startNode) return;
    if (!endNode) { endNode = startNode; endOff = startOff; }
    try {
      const range = document.createRange();
      range.setStart(startNode, startOff);
      range.setEnd(endNode, endOff);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    } catch (e) {}
  }

  /* ─────────────────────────────────
     Save state when picker opens
  ───────────────────────────────── */
  function _openSnapshot() {
    _origHTML     = DOM.editor.innerHTML;
    _selOffsets   = _getSelectionOffsets();
    _hasPreviewed = false;
  }

  /* ─────────────────────────────────
     Live preview — called on every mouseover
  ───────────────────────────────── */
  function _preview(fontValue) {
    if (!_origHTML) return;
    // Restore original HTML so previous preview is undone
    DOM.editor.innerHTML = _origHTML;
    // Re-focus and restore selection
    DOM.editor.focus();
    _restoreSelectionFromOffsets(_selOffsets);
    // Apply preview font only if there is a real selection
    if (_selOffsets && _selOffsets.start !== _selOffsets.end) {
      document.execCommand('fontName', false, fontValue);
      _hasPreviewed = true;
    }
  }

  /* ─────────────────────────────────
     Cancel — revert editor to original
  ───────────────────────────────── */
  function _cancelPreview() {
    if (_origHTML !== null) {
      DOM.editor.innerHTML = _origHTML;
      DOM.editor.focus();
      _restoreSelectionFromOffsets(_selOffsets);
    }
    _clearSnapshot();
    allItems.forEach(i => i.el.classList.remove('is-hovered'));
  }

  /* ─────────────────────────────────
     Commit — re-apply font on clean snapshot so mouseleave revert doesn't
     silently win (mouseleave fires before mousedown when clicking an item)
  ───────────────────────────────── */
  function _commit(fontValue, fontName) {
    if (_origHTML !== null) {
      // Always start from the original clean state, then apply the chosen font.
      // This is necessary because panelEl's mouseleave already reverted the DOM
      // before this mousedown handler fires.
      DOM.editor.innerHTML = _origHTML;
      DOM.editor.focus();
      _restoreSelectionFromOffsets(_selOffsets);
      // Apply font whether it's a selection or just a cursor position
      document.execCommand('fontName', false, fontValue);
    }
    _clearSnapshot();

    currentFont = fontValue;
    currentName = fontName;
    _updateTriggerDisplay(fontValue, fontName);
    allItems.forEach(i => i.el.classList.toggle('is-active', i.font.value === fontValue));

    Storage.autosave();
    StatusBar.show(`Font: ${fontName}`, 'success');
  }

  function _clearSnapshot() {
    _origHTML     = null;
    _selOffsets   = null;
    _hasPreviewed = false;
  }

  /* ─────────────────────────────────
     Build the scrollable font list
  ───────────────────────────────── */
  function _buildList() {
    allItems = [];
    const frag = document.createDocumentFragment();

    FONT_GROUPS.forEach(group => {
      const header = document.createElement('div');
      header.className = 'font-picker__group-label';
      header.textContent = group.label;
      header.setAttribute('aria-hidden', 'true');
      frag.appendChild(header);

      group.fonts.forEach(font => {
        const btn = document.createElement('button');
        btn.className = 'font-picker__item';
        btn.setAttribute('role', 'option');
        btn.setAttribute('aria-label', font.name);
        btn.dataset.fontValue = font.value;
        btn.dataset.fontName  = font.name;

        const nameSpan = document.createElement('span');
        nameSpan.className = 'font-picker__item-name';
        nameSpan.textContent = font.name;
        nameSpan.style.fontFamily = font.value;

        const previewTag = document.createElement('span');
        previewTag.className = 'font-picker__item-preview';
        previewTag.textContent = 'preview';

        btn.appendChild(nameSpan);
        btn.appendChild(previewTag);
        frag.appendChild(btn);
        allItems.push({ el: btn, font });

        if (font.value === currentFont) btn.classList.add('is-active');
      });
    });

    listEl.appendChild(frag);
  }

  /* ─────────────────────────────────
     Wire events
  ───────────────────────────────── */
  function _attachEvents() {
    // Trigger button — mousedown keeps selection alive
    triggerEl.addEventListener('mousedown', (e) => {
      e.preventDefault();
      if (isOpen) {
        _cancelPreview();
        close();
      } else {
        _openSnapshot();
        open();
      }
    });

    // Hover over font item → live preview
    listEl.addEventListener('mouseover', (e) => {
      const item = e.target.closest('.font-picker__item');
      if (!item) return;
      allItems.forEach(i => i.el.classList.remove('is-hovered'));
      item.classList.add('is-hovered');
      _preview(item.dataset.fontValue);
    });

    // Mouse leaves panel → revert to original
    panelEl.addEventListener('mouseleave', () => {
      if (_origHTML !== null) {
        DOM.editor.innerHTML = _origHTML;
        DOM.editor.focus();
        _restoreSelectionFromOffsets(_selOffsets);
        _hasPreviewed = false;
      }
      allItems.forEach(i => i.el.classList.remove('is-hovered'));
    });

    // Click to commit
    listEl.addEventListener('mousedown', (e) => {
      const item = e.target.closest('.font-picker__item');
      if (!item) return;
      e.preventDefault();
      _commit(item.dataset.fontValue, item.dataset.fontName);
      close();
    });

    // Search filter
    searchEl.addEventListener('input', () => {
      _filterFonts(searchEl.value.trim().toLowerCase());
    });

    searchEl.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { _cancelPreview(); close(); }
      if (e.key === 'Enter') {
        const visible = allItems.find(i => i.el.style.display !== 'none');
        if (visible) { _commit(visible.font.value, visible.font.name); close(); }
      }
    });

    // Outside click → cancel (check both trigger wrapper AND panel since panel is on body)
    document.addEventListener('mousedown', (e) => {
      if (!isOpen) return;
      const insideTrigger = $('font-picker').contains(e.target);
      const insidePanel   = panelEl.contains(e.target);
      if (!insideTrigger && !insidePanel) {
        _cancelPreview();
        close();
      }
    });

    // Escape key → cancel
    document.addEventListener('keydown', (e) => {
      if (isOpen && e.key === 'Escape') { _cancelPreview(); close(); }
    });
  }

  /* ─────────────────────────────────
     Filter font list by search query
  ───────────────────────────────── */
  function _filterFonts(query) {
    let anyVisible = false;
    let currentGroupHeader = null;
    let groupHasVisible = false;

    Array.from(listEl.children).forEach(child => {
      if (child.classList.contains('font-picker__group-label')) {
        if (currentGroupHeader) {
          currentGroupHeader.style.display = groupHasVisible ? '' : 'none';
        }
        currentGroupHeader = child;
        groupHasVisible = false;
      } else if (child.classList.contains('font-picker__item')) {
        const name  = child.dataset.fontName.toLowerCase();
        const match = !query || name.includes(query);
        child.style.display = match ? '' : 'none';
        if (match) { groupHasVisible = true; anyVisible = true; }
      }
    });
    if (currentGroupHeader) {
      currentGroupHeader.style.display = groupHasVisible ? '' : 'none';
    }

    // Empty state message
    let emptyEl = listEl.querySelector('.font-picker__empty');
    if (!anyVisible) {
      if (!emptyEl) {
        emptyEl = document.createElement('div');
        emptyEl.className = 'font-picker__empty';
        listEl.appendChild(emptyEl);
      }
      emptyEl.textContent = `No fonts matching "${query}"`;
      emptyEl.style.display = '';
    } else if (emptyEl) {
      emptyEl.style.display = 'none';
    }
  }

  /* ─────────────────────────────────
     Update trigger button appearance
  ───────────────────────────────── */
  function _updateTriggerDisplay(fontValue, fontName) {
    if (!labelEl) return;
    labelEl.textContent      = fontName;
    labelEl.style.fontFamily = fontValue;
  }

  /* ─────────────────────────────────
     PUBLIC: open
  ───────────────────────────────── */
  function open() {
    isOpen = true;

    // Portal positioning — panel lives on <body>, position via fixed coords
    const rect = triggerEl.getBoundingClientRect();
    panelEl.style.top  = `${rect.bottom + 4}px`;
    panelEl.style.left = `${rect.left}px`;

    // Clamp to viewport right edge
    const panelW = 240;
    if (rect.left + panelW > window.innerWidth - 8) {
      panelEl.style.left = `${window.innerWidth - panelW - 8}px`;
    }

    panelEl.removeAttribute('hidden');
    panelEl.setAttribute('aria-hidden', 'false');
    triggerEl.setAttribute('aria-expanded', 'true');
    requestAnimationFrame(() => panelEl.classList.add('is-open'));

    setTimeout(() => {
      searchEl.value = '';
      _filterFonts('');
      searchEl.focus();
      const active = listEl.querySelector('.is-active');
      if (active) active.scrollIntoView({ block: 'nearest' });
    }, 60);
  }

  /* ─────────────────────────────────
     PUBLIC: close
  ───────────────────────────────── */
  function close() {
    isOpen = false;
    panelEl.classList.remove('is-open');
    triggerEl.setAttribute('aria-expanded', 'false');
    allItems.forEach(i => i.el.classList.remove('is-hovered'));
    setTimeout(() => {
      panelEl.setAttribute('hidden', '');
      panelEl.setAttribute('aria-hidden', 'true');
    }, 180);
  }

  /* ─────────────────────────────────
     PUBLIC: sync picker label when cursor moves
  ───────────────────────────────── */
  function syncFromCursor() {
    if (isOpen) return;
    try {
      const raw = document.queryCommandValue('fontName').replace(/['"]/g, '').trim();
      if (!raw) return;

      let matched = null;
      for (const group of FONT_GROUPS) {
        const found = group.fonts.find(f => {
          const first = f.value.split(',')[0].replace(/['"]/g, '').trim();
          return first.toLowerCase() === raw.toLowerCase();
        });
        if (found) { matched = found; break; }
      }

      if (matched && matched.value !== currentFont) {
        currentFont = matched.value;
        currentName = matched.name;
        _updateTriggerDisplay(matched.value, matched.name);
        allItems.forEach(i => {
          i.el.classList.toggle('is-active', i.font.value === matched.value);
        });
      }
    } catch (e) {}
  }

  /* ─────────────────────────────────
     PUBLIC: init
  ───────────────────────────────── */
  function init() {
    triggerEl = $('font-picker-trigger');
    labelEl   = $('font-picker-label');
    panelEl   = $('font-picker-panel');
    listEl    = $('font-picker-list');
    searchEl  = $('font-picker-search');
    if (!triggerEl) return;

    // ── PORTAL: move panel to <body> so it escapes toolbar's overflow:hidden ──
    document.body.appendChild(panelEl);

    _buildList();
    _attachEvents();
    _updateTriggerDisplay(currentFont, currentName);
  }

  return { init, open, close, syncFromCursor };

})();

/* ═══════════════════════════════════════════════════════
   25. INITIALIZATION
═══════════════════════════════════════════════════════ */
function init() {
  // Set creation date
  DOM.metaDate.value = formatDate(new Date());

  // Init subsystems
  ThemeManager.init();
  ZoomManager.init();
  MenuBar.init();
  FontPicker.init();   // custom font picker with live preview

  // Wire up all events
  initToolbar();
  initMenuBar();
  initModals();
  initEditor();
  initKeyboardShortcuts();
  initDragDrop();

  // Load autosaved content
  Storage.loadAutosave();

  // Render recent docs
  UI.renderRecent();

  // Update word count on start
  StatusBar.updateCounts();

  // Announce ready
  StatusBar.show('Ready');

  // Log helpful info
  console.info(
    '%c✦ Quill Editor%c\nLive font preview enabled — hover any font to see it instantly',
    'font-size:16px;font-weight:bold;color:#2d4a8a',
    'font-size:12px;color:#6b6b7a'
  );
}

/* ─────────────────────────────────────────
   Run on DOM ready
───────────────────────────────────────── */
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
