'use strict';

const editorElements = {
  editor: document.getElementById('markdown-input'),
  fileTitleInput: document.getElementById('file-title-input'),
  wordCount: document.getElementById('word-count'),
  charCount: document.getElementById('char-count'),
  fileIndicator: document.getElementById('current-file'),
  statusMessage: document.getElementById('status-message'),
  driveStatus: document.getElementById('drive-status'),
  fileMenu: document.getElementById('file-menu'),
  fileMenuToggle: document.getElementById('file-menu-toggle'),
  fileMenuDropdown: document.getElementById('file-menu-dropdown'),
  tocList: document.getElementById('toc-list'),
  tocEmptyState: document.getElementById('toc-empty'),
  toolbarButtons: document.querySelectorAll('[data-action]'),
  formattingToolbar: document.getElementById('formatting-toolbar'),
  mobileToolbarToggle: document.getElementById('mobile-toolbar-toggle'),
  undoButton: document.querySelector('[data-action="undo"]'),
  redoButton: document.querySelector('[data-action="redo"]'),
  modeToggle: document.getElementById('editor-mode-toggle'),
  dialog: document.getElementById('drive-dialog'),
  dialogClose: document.getElementById('drive-dialog-close'),
  dialogCancel: document.getElementById('drive-dialog-cancel'),
  dialogAlert: document.getElementById('drive-alert'),
  driveFilesWrapper: document.getElementById('drive-files'),
  driveFilesBody: document.getElementById('drive-files-body'),
  driveRefreshButton: document.getElementById('drive-refresh-files'),
  driveDialogTitle: document.getElementById('drive-dialog-title'),
  driveBreadcrumbs: document.getElementById('drive-breadcrumbs'),
  driveFolderUpButton: document.getElementById('drive-folder-up'),
  driveSaveControls: document.getElementById('drive-save-controls'),
  driveFileNameInput: document.getElementById('drive-file-name'),
  driveSaveConfirmButton: document.getElementById('drive-save-confirm'),
  driveOpenButton: document.getElementById('drive-open'),
  driveSaveButton: document.getElementById('drive-save'),
  driveSaveAsButton: document.getElementById('drive-save-as'),
  driveSignInButton: document.getElementById('drive-sign-in'),
  driveSignOutButton: document.getElementById('drive-sign-out'),
  driveConfigStatus: document.getElementById('drive-config-status'),
  driveConfigMessage: document.querySelector('#drive-config-status .cloud-card-message'),
  themeToggle: document.getElementById('theme-toggle'),
  themeToggleIcon: document.querySelector('#theme-toggle .theme-toggle-icon')
};

let currentFileId = null;
let currentFileName = 'Untitled.md';
let pendingFileName = currentFileName;
let isDirty = true;
let gapiReady = false;
let gapiInitPromise = null;
let markdownContent = '';
let htmlContent = '';
let lastNormalizedHtml = '';
let editorMode = 'markdown';
let turndownService = null;
let lastSelection = { start: 0, end: 0 };
let tokenClient = null;
let gisReady = false;
let accessToken = null;
let headerResizeObserver = null;
let isFileMenuOpen = false;

let preferPlainTextRendering = false;
let coarsePointerQuery = null;
let hoverNoneQuery = null;
let touchWidthQuery = null;
let mobileToolbarQuery = null;
let isFormattingToolbarCollapsed = true;

const CONTENT_STORAGE_KEY = 'markdown-editor-content';
const OFFLINE_FILES_STORAGE_KEY = 'markdown-editor-offline-files';
const BASE_DOCUMENT_TITLE = "Mark's Markdown Editor";
const INDENTATION_STRING = '  ';
const PERSISTENCE_DEBOUNCE_MS = 300;
const THEME_STORAGE_KEY = 'markdown-editor-theme';
const THEME_COLOR_MAP = {
  dark: '#1f2937',
  light: '#f8fafc'
};
const COLOR_SCHEME_QUERY = '(prefers-color-scheme: light)';
const TOUCH_RENDER_BREAKPOINT_QUERY = '(max-width: 900px)';
const MOBILE_TOOLBAR_BREAKPOINT_QUERY = '(max-width: 720px)';

let pendingContentPersistence = null;
let persistenceTimeoutId = null;
let persistenceIdleHandle = null;

const HISTORY_LIMIT = 200;
const editorHistory = [];
let historyIndex = -1;
let isNavigatingHistory = false;
const HISTORY_MERGE_INTERVAL = 1200;
let lastHistoryEntryTime = 0;
let lastHistoryWasMergeable = false;

const HTML_VOID_ELEMENTS = new Set([
  'area',
  'base',
  'br',
  'col',
  'embed',
  'hr',
  'img',
  'input',
  'link',
  'meta',
  'param',
  'source',
  'track',
  'wbr'
]);

const DRIVE_ROOT_ID = 'root';
const DRIVE_ROOT_LABEL = 'My Drive';
const VIRTUAL_ROOT_ID = 'files-root';
const VIRTUAL_ROOT_LABEL = 'All files';
const OFFLINE_ROOT_ID = 'offline-root';
const OFFLINE_ROOT_LABEL = 'Offline drafts';

const FILE_SOURCE_VIRTUAL = 'virtual';
const FILE_SOURCE_OFFLINE = 'offline';
const FILE_SOURCE_DRIVE = 'drive';

let driveDialogMode = 'open';
let currentDriveFolderId = VIRTUAL_ROOT_ID;
let currentFolderSource = FILE_SOURCE_VIRTUAL;
let driveFolderPath = [];
let pendingSaveFileId = null;
let pendingSaveFileName = '';
let pendingSaveFileSource = null;
let currentFileSource = null;
let currentOfflineParentId = OFFLINE_ROOT_ID;

const discoveryDocs = ['https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'];
const scopes =
  'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly';

const driveConfigEndpoint = '/config/google-drive.json';

let googleDriveConfig = {
  clientId: '',
  apiKey: ''
};

const defaultMarkdown = `# Welcome to Mark's Markdown Editor

Start typing in the editor to craft your Markdown documents. Use the toolbar buttons to quickly insert Markdown formatting such as **bold**, *italic*, links, lists, tables, and more.

## Features

- Minimal editor focused on Markdown syntax
- Word and character counts update automatically
- Save your documents to Google Drive
- Install the app to work offline as a Progressive Web App

> Tip: Provide Google Drive credentials via your secure runtime configuration (or the \`google-oauth-client-id\` meta tag for local development) to enable Google Drive sync.
`;

function getStoredThemePreference() {
  const stored = localStorage.getItem(THEME_STORAGE_KEY);
  return stored === 'light' || stored === 'dark' ? stored : null;
}

function updateThemeColorMeta(theme) {
  const themeColorMeta = document.querySelector('meta[name="theme-color"]');
  if (!themeColorMeta) {
    return;
  }
  const color = THEME_COLOR_MAP[theme] ?? THEME_COLOR_MAP.dark;
  themeColorMeta.setAttribute('content', color);
}

function updateThemeToggle(theme) {
  if (!editorElements.themeToggle) {
    return;
  }
  const nextLabel = theme === 'light' ? 'Switch to dark mode' : 'Switch to light mode';
  editorElements.themeToggle.setAttribute('aria-pressed', theme === 'light' ? 'true' : 'false');
  editorElements.themeToggle.setAttribute('title', nextLabel);
  editorElements.themeToggle.setAttribute('aria-label', nextLabel);
  editorElements.themeToggle.dataset.theme = theme;

  if (editorElements.themeToggleIcon) {
    editorElements.themeToggleIcon.classList.remove('fa-sun', 'fa-moon');
    editorElements.themeToggleIcon.classList.add(theme === 'light' ? 'fa-sun' : 'fa-moon');
  }
}

function applyThemePreference(theme, { persist = true } = {}) {
  if (theme !== 'light' && theme !== 'dark') {
    return;
  }
  document.documentElement.dataset.theme = theme;
  updateThemeColorMeta(theme);
  updateThemeToggle(theme);
  if (persist) {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  } else {
    localStorage.removeItem(THEME_STORAGE_KEY);
  }
}

function initializeThemePreference() {
  const storedTheme = getStoredThemePreference();
  const prefersLight =
    typeof window !== 'undefined' && window.matchMedia && window.matchMedia(COLOR_SCHEME_QUERY).matches;
  const initialTheme = storedTheme ?? (prefersLight ? 'light' : 'dark');
  applyThemePreference(initialTheme, { persist: Boolean(storedTheme) });

  if (editorElements.themeToggle) {
    editorElements.themeToggle.addEventListener('click', () => {
      const currentTheme = document.documentElement.dataset.theme === 'light' ? 'light' : 'dark';
      const nextTheme = currentTheme === 'light' ? 'dark' : 'light';
      applyThemePreference(nextTheme);
    });
  }

  if (!storedTheme && typeof window !== 'undefined' && window.matchMedia) {
    const mediaQuery = window.matchMedia(COLOR_SCHEME_QUERY);
    const handleChange = (event) => {
      if (!getStoredThemePreference()) {
        applyThemePreference(event.matches ? 'light' : 'dark', { persist: false });
      }
    };
    if (typeof mediaQuery.addEventListener === 'function') {
      mediaQuery.addEventListener('change', handleChange);
    } else if (typeof mediaQuery.addListener === 'function') {
      mediaQuery.addListener(handleChange);
    }
  }
}

function canUndo() {
  return historyIndex > 0;
}

function canRedo() {
  return historyIndex >= 0 && historyIndex < editorHistory.length - 1;
}

function updateUndoRedoButtons() {
  if (editorElements.undoButton) {
    const disabled = !canUndo();
    editorElements.undoButton.disabled = disabled;
    editorElements.undoButton.setAttribute('aria-disabled', disabled ? 'true' : 'false');
  }
  if (editorElements.redoButton) {
    const disabled = !canRedo();
    editorElements.redoButton.disabled = disabled;
    editorElements.redoButton.setAttribute('aria-disabled', disabled ? 'true' : 'false');
  }
}

function saveHistoryEntry(entry, { reset = false, behavior = 'push' } = {}) {
  if (isNavigatingHistory) {
    return;
  }

  const now = Date.now();

  if (reset) {
    editorHistory.length = 0;
    editorHistory.push(entry);
    historyIndex = editorHistory.length ? editorHistory.length - 1 : -1;
    lastHistoryEntryTime = now;
    lastHistoryWasMergeable = false;
    updateUndoRedoButtons();
    return;
  }

  const lastEntry = historyIndex >= 0 ? editorHistory[historyIndex] : null;

  if (lastEntry && lastEntry.content === entry.content) {
    if (
      lastEntry.selectionStart !== entry.selectionStart ||
      lastEntry.selectionEnd !== entry.selectionEnd
    ) {
      editorHistory[historyIndex] = { ...entry };
    }
    lastHistoryEntryTime = now;
    lastHistoryWasMergeable = behavior === 'merge';
    updateUndoRedoButtons();
    return;
  }

  if (
    behavior === 'merge' &&
    lastEntry &&
    lastHistoryWasMergeable &&
    now - lastHistoryEntryTime <= HISTORY_MERGE_INTERVAL
  ) {
    const selectionDelta = entry.selectionStart - lastEntry.selectionStart;
    const contentDelta = entry.content.length - lastEntry.content.length;
    const caretUnchanged =
      entry.selectionStart === entry.selectionEnd &&
      lastEntry.selectionStart === lastEntry.selectionEnd;
    if (caretUnchanged && selectionDelta === contentDelta) {
      editorHistory[historyIndex] = { ...entry };
      lastHistoryEntryTime = now;
      lastHistoryWasMergeable = true;
      updateUndoRedoButtons();
      return;
    }
  }

  if (historyIndex < editorHistory.length - 1) {
    editorHistory.splice(historyIndex + 1);
  }

  editorHistory.push(entry);

  if (editorHistory.length > HISTORY_LIMIT) {
    editorHistory.shift();
  }

  historyIndex = editorHistory.length ? editorHistory.length - 1 : -1;
  lastHistoryEntryTime = now;
  lastHistoryWasMergeable = behavior === 'merge';

  updateUndoRedoButtons();
}

function undo() {
  if (editorMode !== 'markdown') {
    return;
  }

  if (!canUndo()) {
    return;
  }

  historyIndex -= 1;
  const entry = editorHistory[historyIndex];
  if (!entry) {
    historyIndex += 1;
    return;
  }

  isNavigatingHistory = true;
  try {
    applyEditorUpdate(entry.content, entry.selectionStart, entry.selectionEnd, {
      focus: true,
      markDirty: true,
      persistContent: true,
      recordHistory: false
    });
  } finally {
    isNavigatingHistory = false;
    updateUndoRedoButtons();
  }
}

function redo() {
  if (editorMode !== 'markdown') {
    return;
  }

  if (!canRedo()) {
    return;
  }

  historyIndex += 1;
  const entry = editorHistory[historyIndex];
  if (!entry) {
    historyIndex -= 1;
    return;
  }

  isNavigatingHistory = true;
  try {
    applyEditorUpdate(entry.content, entry.selectionStart, entry.selectionEnd, {
      focus: true,
      markDirty: true,
      persistContent: true,
      recordHistory: false
    });
  } finally {
    isNavigatingHistory = false;
    updateUndoRedoButtons();
  }
}

function configureMarkdownConverters() {
  if (window.marked?.setOptions) {
    window.marked.setOptions({
      gfm: true,
      breaks: true,
      headerIds: false,
      mangle: false
    });
  }

  if (!turndownService && window.TurndownService) {
    turndownService = new window.TurndownService({
      headingStyle: 'atx',
      codeBlockStyle: 'fenced',
      bulletListMarker: '-'
    });
    if (window.turndownPluginGfm?.gfm) {
      turndownService.use(window.turndownPluginGfm.gfm);
    }
  }
}

function convertMarkdownToHtml(markdown) {
  if (window.marked?.parse) {
    return window.marked.parse(markdown ?? '').trim();
  }
  return markdown ?? '';
}

function convertHtmlToMarkdown(html) {
  if (turndownService) {
    try {
      return turndownService.turndown(html ?? '');
    } catch (error) {
      console.warn('Unable to convert HTML to Markdown:', error);
    }
  }
  return html ?? '';
}

function formatHtmlContentForEditor(html) {
  if (!html) {
    return '';
  }

  const normalized = html.replace(/\r\n/gu, '\n').replace(/>\s+</gu, '>\n<');
  const lines = normalized.split('\n');
  const formatted = [];
  let indentLevel = 0;
  const indentUnit = '  ';

  lines.forEach((line) => {
    if (!line) {
      return;
    }

    const trimmed = line.trim();
    if (!trimmed) {
      return;
    }

    const isTag = trimmed.startsWith('<');
    if (!isTag) {
      formatted.push(line);
      return;
    }

    const isComment = /^<!--/u.test(trimmed);
    const isClosingTag = /^<\//u.test(trimmed);
    const tagMatch = trimmed.match(/^<([\w:-]+)/u);
    const tagName = tagMatch ? tagMatch[1].toLowerCase() : '';
    const isVoidElement = HTML_VOID_ELEMENTS.has(tagName);
    const isSelfClosing = /\/>$/u.test(trimmed) || isVoidElement;

    if (isClosingTag && !isSelfClosing) {
      indentLevel = Math.max(indentLevel - 1, 0);
    }

    const indentation = indentUnit.repeat(indentLevel);
    formatted.push(`${indentation}${trimmed}`);

    if (
      !isComment &&
      !isClosingTag &&
      !isSelfClosing &&
      !trimmed.includes('</')
    ) {
      indentLevel += 1;
    }
  });

  return formatted.join('\n').replace(/\n+$/u, '');
}

function prepareHtmlContentForConversion(html) {
  if (!html) {
    return '';
  }

  return html
    .replace(/\r\n/gu, '\n')
    .replace(/^[\t ]+(?=<)/gmu, '')
    .replace(/\n[\t ]+(?=<)/gu, '\n')
    .replace(/\n{3,}/gu, '\n\n');
}

function slugifyHeadingText(text) {
  return text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/gu, '')
    .replace(/[^a-z0-9]+/gu, '-')
    .replace(/^-+|-+$/gu, '')
    .replace(/-{2,}/gu, '-');
}

function extractHeadingText(rawHeading) {
  return rawHeading
    .replace(/!\[(.+?)\]\(.*?\)/gu, '$1')
    .replace(/\[(.+?)\]\(.*?\)/gu, '$1')
    .replace(/[`*_~]/gu, '')
    .trim();
}

function createHeadingSpacingMap(markdown) {
  const lines = (markdown || '').split('\n');
  const spacingMap = new Map();

  for (let index = 0; index < lines.length; index += 1) {
    const match = lines[index].match(/^\s{0,3}(#{1,6})\s+(.*)$/u);
    if (!match) {
      continue;
    }

    let blankLines = 0;
    let pointer = index + 1;
    while (pointer < lines.length && lines[pointer].trim() === '') {
      blankLines += 1;
      pointer += 1;
    }

    const level = match[1].length;
    const headingText = extractHeadingText(match[2]);
    const key = `${level}:${headingText.toLowerCase()}`;
    if (!spacingMap.has(key)) {
      spacingMap.set(key, []);
    }
    spacingMap.get(key).push(blankLines);
  }

  return spacingMap;
}

function normalizeMarkdownHeadingSpacing(markdown, spacingMap) {
  if (!markdown || !spacingMap?.size) {
    return markdown;
  }

  const workingMap = new Map();
  spacingMap.forEach((values, key) => {
    workingMap.set(key, values.slice());
  });

  const lines = markdown.split('\n');
  const result = [];
  let index = 0;

  while (index < lines.length) {
    const line = lines[index];
    result.push(line);

    const match = line.match(/^\s{0,3}(#{1,6})\s+(.*)$/u);
    if (match) {
      let blankLines = 0;
      let pointer = index + 1;
      while (pointer < lines.length && lines[pointer].trim() === '') {
        blankLines += 1;
        pointer += 1;
      }

      const level = match[1].length;
      const headingText = extractHeadingText(match[2]);
      const key = `${level}:${headingText.toLowerCase()}`;
      const stored = workingMap.get(key);
      const desiredBlankLines = stored && stored.length ? stored.shift() : null;
      const blanksToKeep =
        desiredBlankLines === null || desiredBlankLines === undefined
          ? blankLines
          : Math.min(blankLines, desiredBlankLines);

      for (let kept = 0; kept < blanksToKeep; kept += 1) {
        result.push('');
      }

      index += 1 + blankLines;
      continue;
    }

    index += 1;
  }

  return result.join('\n');
}

function generateHeadingId(text, level, slugCounts) {
  const baseSlug = slugifyHeadingText(text);
  const fallback = `heading-${level}`;
  const slugKey = (baseSlug || fallback).slice(0, 60).replace(/-+$/gu, '') || fallback;
  const currentCount = slugCounts.get(slugKey) || 0;
  const uniqueSlug = currentCount === 0 ? slugKey : `${slugKey}-${currentCount + 1}`;
  slugCounts.set(slugKey, currentCount + 1);
  return uniqueSlug;
}

function ensureMarkdownExtension(name) {
  const trimmed = (name || '').trim();
  if (!trimmed) {
    return 'Untitled.md';
  }
  return /\.md$/iu.test(trimmed) ? trimmed : `${trimmed}.md`;
}

function normalizeDisplayName(name) {
  const trimmed = (name || '').trim();
  return trimmed || 'Untitled.md';
}

function getDefaultOfflineStore() {
  return {
    version: 1,
    folders: {
      [OFFLINE_ROOT_ID]: {
        id: OFFLINE_ROOT_ID,
        name: OFFLINE_ROOT_LABEL,
        parentId: VIRTUAL_ROOT_ID,
        updated: null
      }
    },
    files: {}
  };
}

function getOfflineStore() {
  const raw = localStorage.getItem(OFFLINE_FILES_STORAGE_KEY);
  if (!raw) {
    return getDefaultOfflineStore();
  }
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') {
      return getDefaultOfflineStore();
    }
    if (!parsed.folders || typeof parsed.folders !== 'object') {
      parsed.folders = {};
    }
    if (!parsed.files || typeof parsed.files !== 'object') {
      parsed.files = {};
    }
    if (!parsed.folders[OFFLINE_ROOT_ID]) {
      parsed.folders[OFFLINE_ROOT_ID] = {
        id: OFFLINE_ROOT_ID,
        name: OFFLINE_ROOT_LABEL,
        parentId: VIRTUAL_ROOT_ID,
        updated: null
      };
    }
    return parsed;
  } catch (error) {
    console.warn('Unable to parse offline files store. Resetting to defaults.', error);
    return getDefaultOfflineStore();
  }
}

function saveOfflineStore(store) {
  try {
    localStorage.setItem(OFFLINE_FILES_STORAGE_KEY, JSON.stringify(store));
  } catch (error) {
    console.warn('Failed to persist offline files store.', error);
  }
}

function generateOfflineId(prefix) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function listOfflineEntries(folderId = OFFLINE_ROOT_ID) {
  const store = getOfflineStore();
  const entries = [];
  Object.values(store.folders).forEach((folder) => {
    if (folder.parentId === folderId && folder.id !== folderId) {
      entries.push({
        id: folder.id,
        name: folder.name,
        mimeType: 'application/vnd.google-apps.folder',
        modifiedTime: folder.updated,
        source: FILE_SOURCE_OFFLINE
      });
    }
  });
  Object.values(store.files).forEach((file) => {
    if (file.parentId === folderId) {
      entries.push({
        id: file.id,
        name: file.name,
        mimeType: 'text/plain',
        modifiedTime: file.updated,
        source: FILE_SOURCE_OFFLINE
      });
    }
  });
  const collator = new Intl.Collator(undefined, { sensitivity: 'base' });
  entries.sort((a, b) => {
    const aIsFolder = a.mimeType === 'application/vnd.google-apps.folder';
    const bIsFolder = b.mimeType === 'application/vnd.google-apps.folder';
    if (aIsFolder !== bIsFolder) {
      return aIsFolder ? -1 : 1;
    }
    return collator.compare(a.name || '', b.name || '');
  });
  return entries;
}

function writeOfflineFile({ fileId = null, fileName, content, folderId = OFFLINE_ROOT_ID } = {}) {
  const store = getOfflineStore();
  const targetFolderId = store.folders[folderId] ? folderId : OFFLINE_ROOT_ID;
  const normalizedName = ensureMarkdownExtension(fileName);
  const id = fileId || generateOfflineId('offline-file');
  const timestamp = new Date().toISOString();
  store.files[id] = {
    id,
    name: normalizedName,
    content,
    parentId: targetFolderId,
    updated: timestamp
  };
  saveOfflineStore(store);
  return {
    id,
    name: normalizedName,
    parentId: targetFolderId,
    updated: timestamp
  };
}

function readOfflineFile(fileId) {
  const store = getOfflineStore();
  return store.files[fileId] || null;
}

function getVirtualRootEntries() {
  const entries = [
    {
      id: OFFLINE_ROOT_ID,
      name: OFFLINE_ROOT_LABEL,
      mimeType: 'application/vnd.google-apps.folder',
      modifiedTime: null,
      source: FILE_SOURCE_OFFLINE
    }
  ];
  if (accessToken) {
    entries.push({
      id: DRIVE_ROOT_ID,
      name: DRIVE_ROOT_LABEL,
      mimeType: 'application/vnd.google-apps.folder',
      modifiedTime: null,
      source: FILE_SOURCE_DRIVE
    });
  }
  return entries;
}

function formatModifiedTime(value) {
  if (!value) {
    return '';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  return date.toLocaleString();
}

function updateHeaderOffset() {
  const root = document.documentElement;
  if (!root) {
    return;
  }

  const header = document.querySelector('header');
  const headerHeight = header ? header.getBoundingClientRect().height : 0;
  root.style.setProperty('--header-offset', `${headerHeight}px`);
}

function setupHeaderOffsetTracking() {
  updateHeaderOffset();

  const header = document.querySelector('header');
  if (header && typeof ResizeObserver !== 'undefined') {
    if (headerResizeObserver) {
      headerResizeObserver.disconnect();
    }
    headerResizeObserver = new ResizeObserver(() => updateHeaderOffset());
    headerResizeObserver.observe(header);
  }

  window.addEventListener('resize', updateHeaderOffset);
  window.addEventListener('orientationchange', updateHeaderOffset);
  window.addEventListener('load', updateHeaderOffset);
}

function getDisplayedFileName() {
  return normalizeDisplayName(pendingFileName);
}

function hasTitleChanges() {
  return normalizeDisplayName(pendingFileName) !== normalizeDisplayName(currentFileName);
}

function getPendingFileNameForSaving() {
  return ensureMarkdownExtension(getDisplayedFileName());
}

function updateTitleInput() {
  if (editorElements.fileTitleInput) {
    editorElements.fileTitleInput.value = normalizeDisplayName(pendingFileName);
  }
}

async function init() {
  configureMarkdownConverters();
  setupTouchEditorOptimizations();
  setupMobileToolbarToggle();
  const savedContent = localStorage.getItem(CONTENT_STORAGE_KEY);
  const initialContent = savedContent ?? defaultMarkdown;
  applyEditorUpdate(initialContent, initialContent.length, initialContent.length, {
    persistContent: false,
    markDirty: false,
    focus: false,
    resetHistory: true
  });
  restoreLastFile();
  updateDriveButtons(false);
  setupHeaderOffsetTracking();
  await loadGoogleDriveConfig();
  if (!isDriveConfigured()) {
    setStatus('Provide Google Drive credentials via your runtime configuration to enable Google Drive sync.', 'error');
  }
  attachEventListeners();
  updateModeToggleState();
  registerServiceWorker();
}

function updateCounts(content) {
  const words = content.trim() ? content.trim().split(/\s+/u).length : 0;
  const characters = content.length;
  editorElements.wordCount.textContent = words;
  editorElements.charCount.textContent = characters;
}

function renderInlineMarkdown(text) {
  const fragment = document.createDocumentFragment();
  const pattern = /\*\*([^*]+)\*\*|\*([^*]+)\*|~~([^~]+)~~/gu;
  let lastIndex = 0;
  let match;

  while ((match = pattern.exec(text)) !== null) {
    const matchStart = match.index;
    if (matchStart > lastIndex) {
      fragment.appendChild(document.createTextNode(text.slice(lastIndex, matchStart)));
    }

    if (match[1] !== undefined) {
      fragment.appendChild(document.createTextNode('**'));
      const strong = document.createElement('span');
      strong.classList.add('md-strong');
      strong.textContent = match[1];
      fragment.appendChild(strong);
      fragment.appendChild(document.createTextNode('**'));
    } else if (match[2] !== undefined) {
      fragment.appendChild(document.createTextNode('*'));
      const emphasis = document.createElement('span');
      emphasis.classList.add('md-em');
      emphasis.textContent = match[2];
      fragment.appendChild(emphasis);
      fragment.appendChild(document.createTextNode('*'));
    } else if (match[3] !== undefined) {
      fragment.appendChild(document.createTextNode('~~'));
      const strike = document.createElement('span');
      strike.classList.add('md-strike');
      strike.textContent = match[3];
      fragment.appendChild(strike);
      fragment.appendChild(document.createTextNode('~~'));
    }

    lastIndex = pattern.lastIndex;
  }

  if (lastIndex < text.length) {
    fragment.appendChild(document.createTextNode(text.slice(lastIndex)));
  }

  if (!fragment.childNodes.length) {
    fragment.appendChild(document.createTextNode(text));
  }

  return fragment;
}

function renderFormattedMarkdown(content) {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  editor.classList.remove('html-mode');
  const tocContainer = document.getElementById('table-of-contents');
  if (tocContainer) {
    tocContainer.hidden = preferPlainTextRendering;
  }
  const lines = content.length ? content.split(/\n/u) : [''];
  const headings = [];
  const slugCounts = new Map();
  if (preferPlainTextRendering) {
    lines.forEach((line) => {
      const headingMatch = line.match(/^(#{1,6})\s+(.*)$/u);
      if (!headingMatch) {
        return;
      }
      const level = Math.min(headingMatch[1].length, 6);
      const headingText = extractHeadingText(headingMatch[2]);
      const displayText = headingText || 'Untitled heading';
      const headingId = generateHeadingId(displayText, level, slugCounts);
      headings.push({ id: headingId, level, text: displayText });
    });
    editor.textContent = content;
    editor.dataset.rendering = 'plain';
    updateTableOfContents(headings);
    return;
  }

  const fragment = document.createDocumentFragment();

  lines.forEach((line, index) => {
    const lineElement = document.createElement('div');
    lineElement.classList.add('editor-line');

    const headingMatch = line.match(/^(#{1,6})\s+(.*)$/u);
    if (headingMatch) {
      const level = Math.min(headingMatch[1].length, 6);
      lineElement.classList.add(`heading-${level}`);
      const headingText = extractHeadingText(headingMatch[2]);
      const displayText = headingText || 'Untitled heading';
      const headingId = generateHeadingId(displayText, level, slugCounts);
      lineElement.id = headingId;
      headings.push({ id: headingId, level, text: displayText });
    }

    if (!line) {
      lineElement.classList.add('is-empty');
      lineElement.innerHTML = '&#8203;';
    } else {
      lineElement.textContent = '';
      lineElement.appendChild(renderInlineMarkdown(line));
    }

    fragment.appendChild(lineElement);

    if (index < lines.length - 1) {
      const lineBreak = document.createElement('span');
      lineBreak.classList.add('editor-line-break');
      lineBreak.setAttribute('aria-hidden', 'true');
      lineBreak.textContent = '\n';
      fragment.appendChild(lineBreak);
    }
  });

  editor.innerHTML = '';
  editor.appendChild(fragment);
  editor.dataset.rendering = 'formatted';
  updateTableOfContents(headings);
}

function renderHtmlEditor(content) {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  editor.classList.add('html-mode');
  editor.textContent = content || '';
  editor.dataset.rendering = 'html';
  const tocContainer = document.getElementById('table-of-contents');
  if (tocContainer) {
    tocContainer.hidden = true;
  }
  if (editorElements.tocList) {
    editorElements.tocList.innerHTML = '';
    editorElements.tocList.hidden = true;
  }
  if (editorElements.tocEmptyState) {
    editorElements.tocEmptyState.hidden = true;
  }
}

function computePlainTextRenderingPreference() {
  const coarse = coarsePointerQuery ? coarsePointerQuery.matches : false;
  const noHover = hoverNoneQuery ? hoverNoneQuery.matches : false;
  const narrow = touchWidthQuery ? touchWidthQuery.matches : false;
  return (coarse || noHover) && narrow;
}

function applyPlainTextRenderingPreference({ forceUpdate = false } = {}) {
  const body = document.body;
  if (!body) {
    return;
  }

  const nextPreference = computePlainTextRenderingPreference();
  if (!forceUpdate && nextPreference === preferPlainTextRendering) {
    return;
  }

  preferPlainTextRendering = nextPreference;
  body.classList.toggle('touch-editor', preferPlainTextRendering);

  if (editorMode !== 'markdown') {
    updateHeaderOffset();
    return;
  }

  const selection = getSelectionOffsets();
  renderFormattedMarkdown(markdownContent);

  if (document.activeElement === editorElements.editor) {
    setSelectionRange(selection.start, selection.end);
  }

  updateHeaderOffset();
}

function setupTouchEditorOptimizations() {
  if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
    preferPlainTextRendering = false;
    return;
  }

  coarsePointerQuery = window.matchMedia('(pointer: coarse)');
  hoverNoneQuery = window.matchMedia('(hover: none)');
  touchWidthQuery = window.matchMedia(TOUCH_RENDER_BREAKPOINT_QUERY);

  const handleChange = () => applyPlainTextRenderingPreference();

  if (typeof coarsePointerQuery.addEventListener === 'function') {
    coarsePointerQuery.addEventListener('change', handleChange);
  } else if (typeof coarsePointerQuery.addListener === 'function') {
    coarsePointerQuery.addListener(handleChange);
  }
  if (typeof hoverNoneQuery.addEventListener === 'function') {
    hoverNoneQuery.addEventListener('change', handleChange);
  } else if (typeof hoverNoneQuery.addListener === 'function') {
    hoverNoneQuery.addListener(handleChange);
  }
  if (typeof touchWidthQuery.addEventListener === 'function') {
    touchWidthQuery.addEventListener('change', handleChange);
  } else if (typeof touchWidthQuery.addListener === 'function') {
    touchWidthQuery.addListener(handleChange);
  }

  applyPlainTextRenderingPreference({ forceUpdate: true });
}

function applyFormattingToolbarVisibility() {
  const toolbar = editorElements.formattingToolbar;
  const toggle = editorElements.mobileToolbarToggle;
  if (!toolbar || !toggle) {
    return;
  }

  const label = toggle.querySelector('.mobile-toolbar-label');
  const isMobile = mobileToolbarQuery ? mobileToolbarQuery.matches : false;

  if (!isMobile) {
    toolbar.hidden = false;
    toggle.hidden = true;
    toggle.setAttribute('aria-expanded', 'true');
    if (label) {
      label.textContent = 'Hide formatting';
    }
    updateHeaderOffset();
    return;
  }

  toggle.hidden = false;
  const expanded = !isFormattingToolbarCollapsed;
  toolbar.hidden = isFormattingToolbarCollapsed;
  toggle.setAttribute('aria-expanded', expanded ? 'true' : 'false');
  if (label) {
    label.textContent = expanded ? 'Hide formatting' : 'Show formatting';
  }
  updateHeaderOffset();
}

function setFormattingToolbarCollapsed(collapsed) {
  if (isFormattingToolbarCollapsed === collapsed) {
    applyFormattingToolbarVisibility();
    return;
  }

  isFormattingToolbarCollapsed = collapsed;
  applyFormattingToolbarVisibility();
}

function setupMobileToolbarToggle() {
  const toolbar = editorElements.formattingToolbar;
  const toggle = editorElements.mobileToolbarToggle;
  if (!toolbar || !toggle) {
    return;
  }

  if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
    toggle.hidden = true;
    toolbar.hidden = false;
    updateHeaderOffset();
    return;
  }

  mobileToolbarQuery = window.matchMedia(MOBILE_TOOLBAR_BREAKPOINT_QUERY);

  toggle.addEventListener('click', () => {
    if (!mobileToolbarQuery || !mobileToolbarQuery.matches) {
      return;
    }
    setFormattingToolbarCollapsed(!isFormattingToolbarCollapsed);
  });

  const handleBreakpointChange = () => {
    if (mobileToolbarQuery.matches) {
      setFormattingToolbarCollapsed(true);
    } else {
      setFormattingToolbarCollapsed(false);
    }
  };

  if (typeof mobileToolbarQuery.addEventListener === 'function') {
    mobileToolbarQuery.addEventListener('change', handleBreakpointChange);
  } else if (typeof mobileToolbarQuery.addListener === 'function') {
    mobileToolbarQuery.addListener(handleBreakpointChange);
  }

  setFormattingToolbarCollapsed(mobileToolbarQuery.matches);
}

function getHtmlEditorContent() {
  const editor = editorElements.editor;
  if (!editor) {
    return '';
  }
  const text = editor.textContent || '';
  return text.replace(/\u200B/gu, '');
}

function setToolbarDisabled(disabled) {
  editorElements.toolbarButtons.forEach((button) => {
    button.disabled = disabled;
    button.setAttribute('aria-disabled', disabled ? 'true' : 'false');
  });
  if (!disabled) {
    updateUndoRedoButtons();
  }
}

function updateModeToggleState() {
  if (!editorElements.modeToggle) {
    return;
  }
  const isHtmlMode = editorMode === 'html';
  const buttonLabel = editorElements.modeToggle.querySelector('.button-label');
  const nextModeText = isHtmlMode ? 'Markdown' : 'HTML';
  const toggleDescription = isHtmlMode ? 'Switch to Markdown mode' : 'Switch to HTML mode';
  if (buttonLabel) {
    buttonLabel.textContent = nextModeText;
  } else {
    editorElements.modeToggle.textContent = `Switch to ${nextModeText}`;
  }
  editorElements.modeToggle.setAttribute('aria-label', toggleDescription);
  editorElements.modeToggle.setAttribute('title', toggleDescription);
  editorElements.modeToggle.setAttribute('aria-pressed', isHtmlMode ? 'true' : 'false');
}

function enterHtmlMode() {
  if (!turndownService) {
    configureMarkdownConverters();
  }
  const convertersReady = Boolean(window.marked?.parse) && Boolean(turndownService);
  const convertedHtml = convertMarkdownToHtml(markdownContent);
  const formattedHtml = convertersReady
    ? formatHtmlContentForEditor(convertedHtml)
    : convertedHtml;
  htmlContent = formattedHtml;
  lastNormalizedHtml = convertersReady
    ? prepareHtmlContentForConversion(formattedHtml)
    : formattedHtml;
  editorMode = 'html';
  renderHtmlEditor(htmlContent);
  setToolbarDisabled(true);
  updateModeToggleState();
  if (editorElements.editor) {
    editorElements.editor.setAttribute('aria-label', 'HTML input');
  }
  lastSelection = { start: 0, end: 0 };
  const statusMessage = convertersReady
    ? 'HTML mode enabled. Edit the generated HTML or switch back to Markdown.'
    : 'HTML conversion libraries unavailable. Editing will use raw Markdown text.';
  setStatus(statusMessage, convertersReady ? 'info' : 'error');
  focusEditor();
  const selection = window.getSelection();
  if (selection && editorElements.editor) {
    const range = document.createRange();
    range.selectNodeContents(editorElements.editor);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  }
}

function exitHtmlMode() {
  const previousMarkdown = markdownContent;
  const latestHtml = getHtmlEditorContent();
  const preparedHtml = prepareHtmlContentForConversion(latestHtml);
  const convertersReady = Boolean(window.marked?.parse) && Boolean(turndownService);
  htmlContent = preparedHtml;

  let nextMarkdown = previousMarkdown;
  let hasChanged = false;

  if (convertersReady) {
    if (preparedHtml !== lastNormalizedHtml) {
      const headingSpacingMap = createHeadingSpacingMap(previousMarkdown);
      const convertedMarkdown = convertHtmlToMarkdown(preparedHtml);
      nextMarkdown = normalizeMarkdownHeadingSpacing(convertedMarkdown, headingSpacingMap);
      hasChanged = nextMarkdown !== previousMarkdown;
    }
  } else {
    nextMarkdown = preparedHtml;
    hasChanged = nextMarkdown !== previousMarkdown;
  }

  editorMode = 'markdown';
  applyEditorUpdate(nextMarkdown, nextMarkdown.length, nextMarkdown.length, {
    focus: true,
    markDirty: hasChanged || isDirty,
    persistContent: true
  });
  setToolbarDisabled(false);
  updateModeToggleState();
  if (editorElements.editor) {
    editorElements.editor.setAttribute('aria-label', 'Markdown input');
  }
  setStatus('Markdown mode enabled.', 'info');
  lastNormalizedHtml = preparedHtml;
}

function toggleEditorMode() {
  if (editorMode === 'markdown') {
    enterHtmlMode();
  } else {
    exitHtmlMode();
  }
}

function updateTableOfContents(headings) {
  const { tocList, tocEmptyState } = editorElements;
  if (!tocList || !tocEmptyState) {
    return;
  }

  tocList.innerHTML = '';

  if (headings.length === 0) {
    tocList.hidden = true;
    tocEmptyState.hidden = false;
    return;
  }

  tocList.hidden = false;
  tocEmptyState.hidden = true;

  const fragment = document.createDocumentFragment();

  headings.forEach((heading) => {
    const item = document.createElement('li');
    item.classList.add(`toc-level-${heading.level}`);

    const link = document.createElement('a');
    link.href = `#${heading.id}`;
    link.textContent = heading.text;
    link.addEventListener('click', () => {
      const target = document.getElementById(heading.id);
      window.requestAnimationFrame(() => {
        if (target) {
          const selection = window.getSelection();
          if (selection) {
            const range = document.createRange();
            range.selectNodeContents(target);
            range.collapse(true);
            selection.removeAllRanges();
            selection.addRange(range);
          }
        }
        focusEditor();
        updateSelectionCache();
      });
    });

    item.appendChild(link);
    fragment.appendChild(item);
  });

  tocList.appendChild(fragment);
}

function getPlainTextFromEditor() {
  const editor = editorElements.editor;
  if (!editor) {
    return '';
  }
  const lineNodes = editor.querySelectorAll('.editor-line');
  if (lineNodes.length === 0) {
    return (editor.textContent || '').replace(/\u200B/gu, '');
  }

  const lines = Array.from(lineNodes, (line) => (line.textContent || '').replace(/\u200B/gu, ''));
  return lines.join('\n');
}

function measureTextLengthToBoundary(container, offset) {
  const editor = editorElements.editor;
  if (!editor) {
    return 0;
  }

  const boundaryRange = document.createRange();
  boundaryRange.selectNodeContents(editor);

  try {
    boundaryRange.setEnd(container, offset);
  } catch (error) {
    console.warn('Unable to measure selection boundary:', error);
    return 0;
  }

  const fragment = boundaryRange.cloneContents();
  const text = fragment.textContent || '';
  return text.replace(/\u200B/gu, '').length;
}

function getSelectionOffsets() {
  const editor = editorElements.editor;
  if (!editor) {
    return { start: 0, end: 0 };
  }

  if (editorMode !== 'markdown') {
    return { start: 0, end: 0 };
  }

  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return { start: 0, end: 0 };
  }

  const range = selection.getRangeAt(0);
  if (!editor.contains(range.startContainer) || !editor.contains(range.endContainer)) {
    return { start: 0, end: 0 };
  }

  const start = measureTextLengthToBoundary(range.startContainer, range.startOffset);
  const end = measureTextLengthToBoundary(range.endContainer, range.endOffset);

  const clampedStart = Math.max(0, Math.min(start, markdownContent.length));
  const clampedEnd = Math.max(0, Math.min(end, markdownContent.length));

  const normalizedStart = Math.min(clampedStart, clampedEnd);
  const normalizedEnd = Math.max(clampedStart, clampedEnd);

  return { start: normalizedStart, end: normalizedEnd };
}

function resolveOffset(offset) {
  const editor = editorElements.editor;
  if (!editor) {
    return { node: null, offset: 0 };
  }

  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null);
  let node = walker.nextNode();
  let traversed = 0;

  while (node) {
    const text = node.textContent || '';
    const clean = text.replace(/\u200B/gu, '');
    const length = clean.length;

    if (traversed + length >= offset) {
      if (length === 0) {
        return { node, offset: 0 };
      }

      const withinNode = offset - traversed;
      let actualOffset = 0;
      let consumed = 0;

      for (let index = 0; index < text.length; index += 1) {
        if (text[index] === '\u200B') {
          continue;
        }

        if (consumed === withinNode) {
          actualOffset = index;
          break;
        }

        consumed += 1;
        actualOffset = index + 1;
      }

      if (withinNode === length) {
        actualOffset = text.length;
      }

      return { node, offset: actualOffset };
    }

    traversed += length;
    node = walker.nextNode();
  }

  return { node: editor, offset: editor.childNodes.length };
}

function findFirstTextNode(node) {
  if (!node) {
    return null;
  }

  if (node.nodeType === Node.TEXT_NODE) {
    return node;
  }

  const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT, null);
  return walker.nextNode();
}

function findNextTextNode(node) {
  let current = node;
  while (current) {
    let sibling = current.nextSibling;
    while (sibling) {
      const textNode = findFirstTextNode(sibling);
      if (textNode) {
        return textNode;
      }
      sibling = sibling.nextSibling;
    }
    current = current.parentNode;
  }
  return null;
}

function normalizeCaretPosition(position) {
  const { node, offset } = position;
  if (!node || node.nodeType !== Node.TEXT_NODE) {
    return position;
  }

  const text = node.textContent || '';
  if (text !== '\n' || offset < text.length) {
    return position;
  }

  const nextTextNode = findNextTextNode(node);
  if (!nextTextNode) {
    return position;
  }

  const nextText = nextTextNode.textContent || '';
  const nextOffset = nextText === '\u200B' ? nextText.length : 0;
  return { node: nextTextNode, offset: nextOffset };
}

function setSelectionRange(start, end) {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  const totalLength = markdownContent.length;
  const clampedStart = Math.max(0, Math.min(start, totalLength));
  const clampedEnd = Math.max(0, Math.min(end, totalLength));

  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const range = document.createRange();
  const startPosition = normalizeCaretPosition(resolveOffset(clampedStart));
  const endPosition = normalizeCaretPosition(resolveOffset(clampedEnd));

  try {
    if (startPosition.node) {
      range.setStart(startPosition.node, startPosition.offset);
    } else {
      range.setStart(editor, 0);
    }

    if (endPosition.node) {
      range.setEnd(endPosition.node, endPosition.offset);
    } else {
      range.setEnd(editor, editor.childNodes.length);
    }
  } catch (error) {
    console.warn('Unable to set selection range:', error);
    return;
  }

  selection.removeAllRanges();
  selection.addRange(range);
}

function focusEditor() {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }
  if (document.activeElement !== editor) {
    editor.focus();
  }
}

function writeContentToLocalStorage(content) {
  try {
    localStorage.setItem(CONTENT_STORAGE_KEY, content);
  } catch (error) {
    console.warn('Failed to persist editor content:', error);
  }
}

function commitPendingPersistence() {
  if (pendingContentPersistence === null) {
    return;
  }
  writeContentToLocalStorage(pendingContentPersistence);
  pendingContentPersistence = null;
}

function persistEditorContent(content, { immediate = false } = {}) {
  if (immediate) {
    if (persistenceTimeoutId) {
      window.clearTimeout(persistenceTimeoutId);
      persistenceTimeoutId = null;
    }
    if (typeof window.cancelIdleCallback === 'function' && persistenceIdleHandle) {
      window.cancelIdleCallback(persistenceIdleHandle);
      persistenceIdleHandle = null;
    }
    pendingContentPersistence = null;
    writeContentToLocalStorage(content);
    return;
  }

  pendingContentPersistence = content;

  if (persistenceTimeoutId) {
    window.clearTimeout(persistenceTimeoutId);
  }

  if (typeof window.cancelIdleCallback === 'function' && persistenceIdleHandle) {
    window.cancelIdleCallback(persistenceIdleHandle);
    persistenceIdleHandle = null;
  }

  const commit = () => {
    persistenceTimeoutId = null;
    persistenceIdleHandle = null;
    commitPendingPersistence();
  };

  persistenceTimeoutId = window.setTimeout(commit, PERSISTENCE_DEBOUNCE_MS);

  if (typeof window.requestIdleCallback === 'function') {
    persistenceIdleHandle = window.requestIdleCallback(commit, {
      timeout: PERSISTENCE_DEBOUNCE_MS
    });
  }
}

function flushPendingContentPersistence() {
  if (persistenceTimeoutId) {
    window.clearTimeout(persistenceTimeoutId);
    persistenceTimeoutId = null;
  }

  if (typeof window.cancelIdleCallback === 'function' && persistenceIdleHandle) {
    window.cancelIdleCallback(persistenceIdleHandle);
    persistenceIdleHandle = null;
  }

  commitPendingPersistence();
}

function applyEditorUpdate(content, selectionStart = content.length, selectionEnd = selectionStart, options = {}) {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  const {
    markDirty = true,
    persistContent = true,
    focus = true,
    recordHistory = true,
    resetHistory = false,
    historyBehavior = 'push'
  } = options;
  const previousScrollTop = editor.scrollTop;
  const previousScrollLeft = editor.scrollLeft;

  editorMode = 'markdown';
  markdownContent = content;
  renderFormattedMarkdown(content);

  editor.scrollTop = previousScrollTop;
  editor.scrollLeft = previousScrollLeft;

  updateCounts(content);

  if (persistContent) {
    persistEditorContent(content);
  }

  if (markDirty) {
    isDirty = true;
  } else {
    isDirty = false;
  }

  updateFileIndicator();

  if (focus) {
    editor.focus();
  }

  setSelectionRange(selectionStart, selectionEnd);
  lastSelection = { start: selectionStart, end: selectionEnd };

  if (resetHistory) {
    saveHistoryEntry({ content, selectionStart, selectionEnd }, { reset: true });
  } else if (recordHistory) {
    saveHistoryEntry({ content, selectionStart, selectionEnd }, { behavior: historyBehavior });
  } else if (!isNavigatingHistory) {
    updateUndoRedoButtons();
  }
}

function handleEditorInput() {
  if (editorMode === 'html') {
    const htmlValue = getHtmlEditorContent();
    htmlContent = htmlValue;
    const markdownValue = convertHtmlToMarkdown(htmlValue);
    markdownContent = markdownValue;
    updateCounts(markdownValue);
    persistEditorContent(markdownValue);
    isDirty = true;
    updateFileIndicator();
    return;
  }

  const { start, end } = getSelectionOffsets();
  const value = getPlainTextFromEditor();
  applyEditorUpdate(value, start, end, { focus: false, historyBehavior: 'merge' });
}

function handleTitleInputChange() {
  if (!editorElements.fileTitleInput) {
    return;
  }
  pendingFileName = editorElements.fileTitleInput.value;
  updateFileIndicator();
}

function handleTitleInputBlur() {
  if (!editorElements.fileTitleInput) {
    return;
  }
  pendingFileName = normalizeDisplayName(editorElements.fileTitleInput.value);
  updateTitleInput();
  updateFileIndicator();
}

function getSelectedLines(start, end) {
  const contentLength = markdownContent.length;
  const safeStart = Math.max(0, Math.min(start, contentLength));
  const safeEnd = Math.max(0, Math.min(end, contentLength));

  let lineStart = safeStart;
  while (lineStart > 0 && markdownContent[lineStart - 1] !== '\n') {
    lineStart -= 1;
  }

  let lineEnd = safeEnd;
  if (
    lineEnd > lineStart &&
    safeEnd > safeStart &&
    safeEnd > 0 &&
    markdownContent[safeEnd - 1] === '\n'
  ) {
    lineEnd -= 1;
  }

  if (lineEnd < lineStart) {
    lineEnd = lineStart;
  }

  while (lineEnd < contentLength && markdownContent[lineEnd] !== '\n') {
    lineEnd += 1;
  }

  const segment = markdownContent.slice(lineStart, lineEnd);
  const lines = segment ? segment.split('\n') : [''];

  return { lineStart, lineEnd, safeStart, safeEnd, lines };
}

function adjustIndexForAddition(index, lineStart, amount) {
  if (index < lineStart) {
    return index;
  }
  return index + amount;
}

function adjustIndexForRemoval(index, lineStart, amount) {
  if (index <= lineStart) {
    return index;
  }
  if (index <= lineStart + amount) {
    return lineStart;
  }
  return index - amount;
}

function removeIndentationFromLine(line) {
  if (!line) {
    return { line: '', removed: 0 };
  }
  if (line.startsWith('\t')) {
    return { line: line.slice(1), removed: 1 };
  }
  if (line.startsWith(INDENTATION_STRING)) {
    return { line: line.slice(INDENTATION_STRING.length), removed: INDENTATION_STRING.length };
  }
  const match = line.match(/^ +/u);
  if (match) {
    const spacesToRemove = Math.min(match[0].length, INDENTATION_STRING.length);
    return { line: line.slice(spacesToRemove), removed: spacesToRemove };
  }
  return { line, removed: 0 };
}

function handleEditorKeyDown(event) {
  if (event.isComposing) {
    return;
  }

  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  if (editorMode !== 'markdown') {
    return;
  }

  if (event.key === 'Enter') {
    event.preventDefault();

    const { start, end } = getSelectionOffsets();
    const before = markdownContent.slice(0, start);
    const after = markdownContent.slice(end);
    const nextContent = `${before}\n${after}`;
    const caretPosition = start + 1;

    applyEditorUpdate(nextContent, caretPosition, caretPosition);
    return;
  }

  if (event.key === 'Tab') {
    event.preventDefault();

    const { start, end } = getSelectionOffsets();
    const { lineStart, lineEnd, safeStart, safeEnd, lines } = getSelectedLines(start, end);

    if (!event.shiftKey) {
      let adjustedStart = safeStart;
      let adjustedEnd = safeEnd;
      let processed = 0;

      const updatedLines = lines.map((line, index) => {
        const absoluteLineStart = lineStart + processed;
        adjustedStart = adjustIndexForAddition(adjustedStart, absoluteLineStart, INDENTATION_STRING.length);
        adjustedEnd = adjustIndexForAddition(adjustedEnd, absoluteLineStart, INDENTATION_STRING.length);
        processed += line.length;
        if (index < lines.length - 1) {
          processed += 1;
        }
        return `${INDENTATION_STRING}${line}`;
      });

      const nextContent = `${markdownContent.slice(0, lineStart)}${updatedLines.join('\n')}${markdownContent.slice(
        lineEnd
      )}`;

      applyEditorUpdate(nextContent, adjustedStart, adjustedEnd);
      return;
    }

    let adjustedStart = safeStart;
    let adjustedEnd = safeEnd;
    let processed = 0;
    let hasOutdentChange = false;

    const updatedLines = lines.map((line, index) => {
      const absoluteLineStart = lineStart + processed;
      const { line: trimmedLine, removed } = removeIndentationFromLine(line);
      if (removed > 0) {
        adjustedStart = adjustIndexForRemoval(adjustedStart, absoluteLineStart, removed);
        adjustedEnd = adjustIndexForRemoval(adjustedEnd, absoluteLineStart, removed);
        hasOutdentChange = true;
      }
      processed += line.length;
      if (index < lines.length - 1) {
        processed += 1;
      }
      return trimmedLine;
    });

    if (!hasOutdentChange) {
      return;
    }

    const nextContent = `${markdownContent.slice(0, lineStart)}${updatedLines.join('\n')}${markdownContent.slice(
      lineEnd
    )}`;

    applyEditorUpdate(nextContent, adjustedStart, adjustedEnd);
    return;
  }

  if (event.key === 'Backspace') {
    const { start, end } = getSelectionOffsets();
    if (start === 0 && end === 0) {
      return;
    }

    event.preventDefault();

    const deletionStart = start === end ? Math.max(0, start - 1) : start;
    const deletionEnd = end;
    const nextContent = `${markdownContent.slice(0, deletionStart)}${markdownContent.slice(deletionEnd)}`;
    const caretPosition = deletionStart;

    applyEditorUpdate(nextContent, caretPosition, caretPosition);
    return;
  }

  if (event.key === 'Delete') {
    const { start, end } = getSelectionOffsets();
    if (start === markdownContent.length && end === markdownContent.length) {
      return;
    }

    event.preventDefault();

    const deletionStart = start;
    const deletionEnd = start === end ? Math.min(markdownContent.length, end + 1) : end;
    if (deletionStart === deletionEnd) {
      return;
    }

    const nextContent = `${markdownContent.slice(0, deletionStart)}${markdownContent.slice(deletionEnd)}`;

    applyEditorUpdate(nextContent, deletionStart, deletionStart);
  }
}

function updateSelectionCache() {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }
  if (editorMode !== 'markdown') {
    return;
  }
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return;
  }
  const anchor = selection.anchorNode;
  const focus = selection.focusNode;
  if (anchor && focus && editor.contains(anchor) && editor.contains(focus)) {
    lastSelection = getSelectionOffsets();
  }
}

function focusFirstFileMenuItem() {
  if (!editorElements.fileMenuDropdown) {
    return;
  }
  const firstAction = editorElements.fileMenuDropdown.querySelector('button:not(:disabled)');
  firstAction?.focus();
}

function setFileMenuOpen(open, { focusFirst = false, returnFocus = false } = {}) {
  const { fileMenu, fileMenuToggle, fileMenuDropdown } = editorElements;
  if (!fileMenu || !fileMenuToggle || !fileMenuDropdown) {
    return;
  }

  isFileMenuOpen = open;
  fileMenu.dataset.open = open ? 'true' : 'false';
  fileMenuToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
  fileMenuDropdown.dataset.open = open ? 'true' : 'false';

  if (open) {
    fileMenuDropdown.removeAttribute('hidden');
    if (focusFirst) {
      focusFirstFileMenuItem();
    }
  } else {
    fileMenuDropdown.setAttribute('hidden', '');
    if (returnFocus) {
      fileMenuToggle.focus();
    }
  }
}

function toggleFileMenu(options = {}) {
  setFileMenuOpen(!isFileMenuOpen, options);
}

function setStatus(message, type = 'info') {
  editorElements.statusMessage.textContent = message;
  editorElements.statusMessage.className = '';
  if (!message) {
    return;
  }
  if (type === 'error') {
    editorElements.statusMessage.classList.add('alert');
  } else if (type === 'success') {
    editorElements.statusMessage.classList.add('status-success');
  }
}

function updateFileIndicator() {
  const displayName = getDisplayedFileName();
  const hasUnsavedChanges = isDirty || hasTitleChanges();
  const indicator = `${displayName}${hasUnsavedChanges ? ' • Unsaved changes' : ''}`;
  editorElements.fileIndicator.textContent = indicator;
  updateDocumentTitle(displayName, hasUnsavedChanges);
}

function updateDocumentTitle(displayName, hasUnsavedChanges) {
  const safeName = displayName || 'Untitled.md';
  const titleParts = [safeName, BASE_DOCUMENT_TITLE];
  let title = titleParts.join(' – ');
  if (hasUnsavedChanges) {
    title += ' • Unsaved changes';
  }
  document.title = title;
}

function attachEventListeners() {
  const editor = editorElements.editor;
  editor.addEventListener('input', () => handleEditorInput());
  editor.addEventListener('keydown', handleEditorKeyDown);
  editor.addEventListener('keyup', () => updateSelectionCache());
  editor.addEventListener('mouseup', () => updateSelectionCache());
  editor.addEventListener('blur', () => updateSelectionCache());

  if (editorElements.fileTitleInput) {
    editorElements.fileTitleInput.addEventListener('input', () => handleTitleInputChange());
    editorElements.fileTitleInput.addEventListener('blur', () => handleTitleInputBlur());
    editorElements.fileTitleInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        editorElements.fileTitleInput.blur();
      }
    });
  }

  document.addEventListener('selectionchange', () => updateSelectionCache());

  editorElements.toolbarButtons.forEach((button) => {
    button.addEventListener('mousedown', (event) => event.preventDefault());
    button.addEventListener('click', () => applyMarkdown(button.dataset.action));
  });

  if (editorElements.modeToggle) {
    editorElements.modeToggle.addEventListener('mousedown', (event) => event.preventDefault());
    editorElements.modeToggle.addEventListener('click', () => toggleEditorMode());
  }

  if (editorElements.fileMenu && editorElements.fileMenuToggle && editorElements.fileMenuDropdown) {
    editorElements.fileMenuToggle.addEventListener('click', (event) => {
      const focusFirst = !isFileMenuOpen && event.detail === 0;
      toggleFileMenu({ focusFirst });
    });

    editorElements.fileMenuToggle.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        setFileMenuOpen(true, { focusFirst: true });
      } else if (event.key === 'Escape') {
        setFileMenuOpen(false);
      }
    });

    editorElements.fileMenuDropdown.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        setFileMenuOpen(false, { returnFocus: true });
      }
    });

    editorElements.fileMenuDropdown.addEventListener('click', (event) => {
      const button = event.target.closest('button');
      if (!button || button.disabled) {
        return;
      }
      setFileMenuOpen(false);
    });
  }

  editorElements.dialogClose.addEventListener('click', () => closeDialog());
  editorElements.dialogCancel.addEventListener('click', () => closeDialog());

  if (editorElements.driveRefreshButton) {
    editorElements.driveRefreshButton.addEventListener('click', () => {
      refreshDriveFileList();
    });
  }

  if (editorElements.driveFolderUpButton) {
    editorElements.driveFolderUpButton.addEventListener('click', () => {
      navigateToParentFolder();
    });
  }

  editorElements.driveOpenButton.addEventListener('click', () => {
    openDialog('open');
  });

  editorElements.driveSaveButton.addEventListener('click', () => {
    handleSaveButtonClick();
  });

  editorElements.driveSaveAsButton.addEventListener('click', () => {
    openDialog('save');
  });

  if (editorElements.driveSaveConfirmButton) {
    editorElements.driveSaveConfirmButton.addEventListener('click', () => {
      handleDriveSaveConfirm();
    });
  }

  if (editorElements.driveFileNameInput) {
    editorElements.driveFileNameInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        handleDriveSaveConfirm();
      }
    });
    editorElements.driveFileNameInput.addEventListener('input', () => {
      if (!pendingSaveFileId) {
        return;
      }
      const trimmed = editorElements.driveFileNameInput.value.trim();
      if (trimmed !== pendingSaveFileName) {
        clearDriveSelection();
      }
    });
  }

  editorElements.driveSignInButton.addEventListener('click', () => {
    signInToGoogle();
  });

  editorElements.driveSignOutButton.addEventListener('click', () => {
    signOutOfGoogle();
  });

  editorElements.dialog.addEventListener('click', (event) => {
    if (event.target === editorElements.dialog) {
      closeDialog();
    }
  });

  document.addEventListener('click', (event) => {
    if (!isFileMenuOpen) {
      return;
    }
    if (
      !editorElements.fileMenu?.contains(event.target) &&
      !editorElements.fileMenuDropdown?.contains(event.target)
    ) {
      setFileMenuOpen(false);
    }
  });

  document.addEventListener('keydown', (event) => {
    if (!isFileMenuOpen) {
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      setFileMenuOpen(false, { returnFocus: true });
    }
  });
}

function registerServiceWorker() {
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker
        .register('service-worker.js')
        .catch((error) => console.warn('Service worker registration failed:', error));
    });
  }
}

function applyMarkdown(action) {
  if (editorMode !== 'markdown') {
    setStatus('Switch to Markdown mode to use formatting tools.', 'info');
    return;
  }

  focusEditor();
  setSelectionRange(lastSelection.start, lastSelection.end);

  switch (action) {
    case 'undo':
      undo();
      break;
    case 'redo':
      redo();
      break;
    case 'bold':
      wrapSelection('**', '**', 'bold text');
      break;
    case 'italic':
      wrapSelection('*', '*', 'italic text');
      break;
    case 'strikethrough':
      wrapSelection('~~', '~~', 'strikethrough text');
      break;
    case 'inline-code':
      wrapSelection('`', '`', 'code');
      break;
    case 'heading-1':
      applyLinePrefix('# ', 'Heading 1');
      break;
    case 'heading-2':
      applyLinePrefix('## ', 'Heading 2');
      break;
    case 'heading-3':
      applyLinePrefix('### ', 'Heading 3');
      break;
    case 'blockquote':
      applyBlockQuote();
      break;
    case 'link':
      insertLink();
      break;
    case 'image':
      insertImage();
      break;
    case 'ordered-list':
      applyList(true);
      break;
    case 'unordered-list':
      applyList(false);
      break;
    case 'horizontal-rule':
      insertSnippet('\n\n---\n\n');
      break;
    case 'table':
      insertSnippet('\n\n| Column 1 | Column 2 |\n| --- | --- |\n| Item 1 | Item 2 |\n\n');
      break;
    case 'code-block':
      insertCodeBlock();
      break;
    default:
      break;
  }
}

function wrapSelection(before, after, placeholder) {
  const { start, end } = getSelectionOffsets();
  const value = markdownContent;
  const selected = value.slice(start, end) || placeholder;
  const newValue = `${value.slice(0, start)}${before}${selected}${after}${value.slice(end)}`;
  const newStart = start + before.length;
  const newEnd = newStart + selected.length;
  applyEditorUpdate(newValue, newStart, newEnd);
}

function insertSnippet(snippet) {
  const { start, end } = getSelectionOffsets();
  const value = markdownContent;
  const newValue = `${value.slice(0, start)}${snippet}${value.slice(end)}`;
  const cursorPosition = start + snippet.length;
  applyEditorUpdate(newValue, cursorPosition, cursorPosition);
}

function insertCodeBlock() {
  const { start, end } = getSelectionOffsets();
  const value = markdownContent;
  const selected = value.slice(start, end);
  const placeholder = 'code';
  const blockStart = '\n\n```\n';
  const blockEnd = '\n```\n\n';
  const content = selected || placeholder;
  const newValue = `${value.slice(0, start)}${blockStart}${content}${blockEnd}${value.slice(end)}`;
  const selectionStart = start + blockStart.length;
  const selectionEnd = selectionStart + content.length;
  applyEditorUpdate(newValue, selectionStart, selectionEnd);
}

function applyLinePrefix(prefix, placeholder = '') {
  const value = markdownContent;
  const { start, end } = getSelectionOffsets();
  const lineStart = value.lastIndexOf('\n', start - 1) + 1;
  let lineEnd = value.indexOf('\n', end);
  if (lineEnd === -1) {
    lineEnd = value.length;
  }
  const selected = value.slice(lineStart, lineEnd);
  const lines = selected.split('\n');
  const updatedLines = lines.map((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      return placeholder ? `${prefix}${placeholder}` : prefix;
    }
    if (line.startsWith(prefix)) {
      const withoutPrefix = line.slice(prefix.length).trim();
      if (!withoutPrefix && placeholder) {
        return `${prefix}${placeholder}`;
      }
      return line;
    }
    const cleaned = line.replace(/^#{1,6}\s+/u, '').trimStart();
    if (!cleaned && placeholder) {
      return `${prefix}${placeholder}`;
    }
    return `${prefix}${cleaned}`;
  });
  const updated = updatedLines.join('\n');
  const firstLine = updatedLines[0] ?? '';
  const selectionStart = lineStart + Math.min(prefix.length, firstLine.length);
  const selectionEnd = selectionStart + Math.max(firstLine.length - prefix.length, 0);
  const newValue = `${value.slice(0, lineStart)}${updated}${value.slice(lineEnd)}`;
  applyEditorUpdate(newValue, selectionStart, selectionEnd);
}

function applyList(ordered) {
  const value = markdownContent;
  const { start, end } = getSelectionOffsets();
  const lineStart = value.lastIndexOf('\n', start - 1) + 1;
  let lineEnd = value.indexOf('\n', end);
  if (lineEnd === -1) {
    lineEnd = value.length;
  }
  const selected = value.slice(lineStart, lineEnd);
  const lines = selected.split('\n');
  let counter = 1;
  const updatedLines = lines.map((line) => {
    if (!line.trim()) {
      return line;
    }
    if (ordered) {
      const cleaned = line.replace(/^\d+\.\s+/u, '');
      return `${counter++}. ${cleaned}`;
    }
    const cleaned = line.replace(/^[-*+]\s+/u, '');
    return `- ${cleaned}`;
  });
  const updated = updatedLines.join('\n');
  const newValue = `${value.slice(0, lineStart)}${updated}${value.slice(lineEnd)}`;
  applyEditorUpdate(newValue, lineStart, lineStart + updated.length);
}

function applyBlockQuote() {
  const value = markdownContent;
  const { start, end } = getSelectionOffsets();
  const lineStart = value.lastIndexOf('\n', start - 1) + 1;
  let lineEnd = value.indexOf('\n', end);
  if (lineEnd === -1) {
    lineEnd = value.length;
  }
  const selected = value.slice(lineStart, lineEnd);
  const lines = selected.split('\n');
  const hasContent = lines.some((line) => line.trim());
  let placeholderInserted = false;
  const updatedLines = lines.map((line) => {
    const leadingWhitespaceMatch = line.match(/^\s*/u);
    const leadingWhitespace = leadingWhitespaceMatch ? leadingWhitespaceMatch[0] : '';
    const trimmedStart = line.trimStart();
    if (!trimmedStart) {
      if (!hasContent && !placeholderInserted) {
        placeholderInserted = true;
        return `${leadingWhitespace}> Quote text`;
      }
      return `${leadingWhitespace}> `;
    }
    if (trimmedStart.startsWith('>')) {
      return `${leadingWhitespace}${trimmedStart}`;
    }
    return `${leadingWhitespace}> ${trimmedStart}`;
  });
  const updated = updatedLines.join('\n');
  const newValue = `${value.slice(0, lineStart)}${updated}${value.slice(lineEnd)}`;
  const selectionStart = lineStart;
  const selectionEnd = selectionStart + updated.length;
  applyEditorUpdate(newValue, selectionStart, selectionEnd);
}

function insertLink() {
  const url = window.prompt('Enter the URL');
  if (!url) {
    return;
  }
  wrapSelection('[', `](${url})`, 'link text');
}

function insertImage() {
  const url = window.prompt('Enter the image URL');
  if (!url) {
    return;
  }
  wrapSelection('![', `](${url})`, 'alt text');
}

function openDialog(mode = 'open') {
  driveDialogMode = mode;
  clearDriveError();
  updateDriveConfigMessage();
  editorElements.dialog.classList.add('active');
  editorElements.dialog.setAttribute('aria-hidden', 'false');
  if (editorElements.driveFilesWrapper) {
    editorElements.driveFilesWrapper.hidden = true;
  }
  driveFolderPath = [{ id: VIRTUAL_ROOT_ID, name: VIRTUAL_ROOT_LABEL, source: FILE_SOURCE_VIRTUAL }];
  currentDriveFolderId = VIRTUAL_ROOT_ID;
  currentFolderSource = FILE_SOURCE_VIRTUAL;
  pendingSaveFileId = null;
  pendingSaveFileName = '';
  pendingSaveFileSource = null;
  updateDriveDialogMode();
  refreshDriveFileList({ folderId: currentDriveFolderId, source: currentFolderSource });
  if (driveDialogMode === 'save') {
    const defaultName = getPendingFileNameForSaving();
    if (editorElements.driveFileNameInput) {
      editorElements.driveFileNameInput.value = defaultName;
      pendingSaveFileName = defaultName;
      window.setTimeout(() => {
        editorElements.driveFileNameInput?.focus();
        editorElements.driveFileNameInput?.select();
      }, 0);
    }
  } else {
    window.setTimeout(() => {
      editorElements.driveRefreshButton?.focus();
    }, 0);
  }
}

function closeDialog() {
  editorElements.dialog.classList.remove('active');
  editorElements.dialog.setAttribute('aria-hidden', 'true');
  pendingSaveFileId = null;
  pendingSaveFileName = '';
  pendingSaveFileSource = null;
}

function updateDriveDialogMode() {
  const isSaveMode = driveDialogMode === 'save';
  if (editorElements.driveDialogTitle) {
    editorElements.driveDialogTitle.textContent = isSaveMode ? 'Save file' : 'Open file';
  }
  if (editorElements.driveSaveControls) {
    editorElements.driveSaveControls.hidden = !isSaveMode;
  }
  if (!isSaveMode && editorElements.driveFileNameInput) {
    editorElements.driveFileNameInput.value = '';
  }
  updateDrivePathDisplay();
  updateDriveFolderUpButton();
}

function updateDrivePathDisplay() {
  if (!editorElements.driveBreadcrumbs) {
    return;
  }
  if (!driveFolderPath.length) {
    driveFolderPath = [{ id: VIRTUAL_ROOT_ID, name: VIRTUAL_ROOT_LABEL, source: FILE_SOURCE_VIRTUAL }];
  }
  editorElements.driveBreadcrumbs.innerHTML = '';
  driveFolderPath.forEach((entry, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = entry.name || 'Untitled';
    const isLast = index === driveFolderPath.length - 1;
    button.disabled = isLast;
    if (!button.disabled) {
      button.addEventListener('click', () => {
        driveFolderPath = driveFolderPath.slice(0, index + 1);
        const target = driveFolderPath[driveFolderPath.length - 1];
        currentDriveFolderId = target.id;
        currentFolderSource = target.source ?? FILE_SOURCE_VIRTUAL;
        clearDriveSelection();
        refreshDriveFileList({ folderId: target.id, source: currentFolderSource });
      });
    }
    editorElements.driveBreadcrumbs.appendChild(button);
    if (index < driveFolderPath.length - 1) {
      const separator = document.createElement('span');
      separator.classList.add('separator');
      separator.textContent = '/';
      editorElements.driveBreadcrumbs.appendChild(separator);
    }
  });
  updateDriveFolderUpButton();
}

function updateDriveFolderUpButton() {
  if (!editorElements.driveFolderUpButton) {
    return;
  }
  editorElements.driveFolderUpButton.disabled = driveFolderPath.length <= 1;
}

function navigateToParentFolder() {
  if (driveFolderPath.length <= 1) {
    return;
  }
  driveFolderPath = driveFolderPath.slice(0, -1);
  const parent = driveFolderPath[driveFolderPath.length - 1];
  currentDriveFolderId = parent.id;
  currentFolderSource = parent.source ?? FILE_SOURCE_VIRTUAL;
  clearDriveSelection();
  refreshDriveFileList({ folderId: parent.id, source: currentFolderSource });
}

function enterDriveFolder(folder) {
  if (!folder?.id) {
    return;
  }
  const name = folder.name || 'Untitled folder';
  driveFolderPath = [...driveFolderPath, { id: folder.id, name, source: folder.source ?? currentFolderSource }];
  currentDriveFolderId = folder.id;
  currentFolderSource = folder.source ?? currentFolderSource;
  clearDriveSelection();
  refreshDriveFileList({ folderId: folder.id, source: currentFolderSource });
}

function clearDriveSelection() {
  pendingSaveFileId = null;
  pendingSaveFileName = '';
  pendingSaveFileSource = null;
  if (!editorElements.driveFilesBody) {
    return;
  }
  editorElements.driveFilesBody.querySelectorAll('.selected').forEach((row) => {
    row.classList.remove('selected');
  });
}

function selectDriveFileForSave(file, row) {
  if (!file?.id) {
    return;
  }
  clearDriveSelection();
  pendingSaveFileId = file.id;
  pendingSaveFileSource = file.source ?? currentFolderSource;
  if (row) {
    row.classList.add('selected');
  }
  if (editorElements.driveFileNameInput) {
    editorElements.driveFileNameInput.value = ensureMarkdownExtension(file.name || 'Untitled.md');
    pendingSaveFileName = editorElements.driveFileNameInput.value.trim();
    editorElements.driveFileNameInput.focus();
    editorElements.driveFileNameInput.select();
  } else {
    pendingSaveFileName = ensureMarkdownExtension(file.name || 'Untitled.md');
  }
}

async function handleDriveSaveConfirm() {
  if (driveDialogMode !== 'save') {
    return;
  }
  clearDriveError();
  const input = editorElements.driveFileNameInput?.value ?? '';
  if (!input.trim()) {
    showDriveError('Enter a file name to save.');
    editorElements.driveFileNameInput?.focus();
    return;
  }
  const fileName = ensureMarkdownExtension(input);
  const confirmButton = editorElements.driveSaveConfirmButton;
  if (confirmButton) {
    confirmButton.disabled = true;
  }
  let success = false;
  let shouldCloseDialog = false;
  if (currentFolderSource === FILE_SOURCE_OFFLINE) {
    const targetId = pendingSaveFileSource === FILE_SOURCE_OFFLINE ? pendingSaveFileId : null;
    const saved = saveToOffline({ fileId: targetId, fileName, folderId: currentDriveFolderId });
    success = Boolean(saved);
    shouldCloseDialog = success;
  } else if (currentFolderSource === FILE_SOURCE_DRIVE) {
    const options = {
      folderId: currentDriveFolderId,
      fileName
    };
    if (pendingSaveFileId && pendingSaveFileSource === FILE_SOURCE_DRIVE) {
      options.fileId = pendingSaveFileId;
    } else {
      options.forceNew = true;
    }
    success = await saveToDrive(options);
    shouldCloseDialog = success;
  } else {
    showDriveError('Select a location to save the file.');
  }
  if (confirmButton) {
    confirmButton.disabled = false;
  }
  if (success && shouldCloseDialog) {
    closeDialog();
  }
}

function showDriveError(error) {
  if (!error) {
    return;
  }
  if (error?.message === 'popup_closed_by_user' || error?.message === 'user_cancelled') {
    setStatus('Google sign-in was canceled.');
    return;
  }
  console.error(error);
  let message = 'Unable to complete the Google Drive request.';
  const code = error?.result?.error?.code ?? error?.status;
  if (code === 401) {
    clearAccessToken();
    message = 'Google Drive authorization expired. Please sign in again.';
  } else if (typeof error === 'string') {
    message = error;
  } else if (error?.result?.error?.message) {
    message = error.result.error.message;
  } else if (error?.message) {
    message = error.message;
  }
  if (editorElements.dialogAlert) {
    editorElements.dialogAlert.textContent = message;
    editorElements.dialogAlert.hidden = false;
  }
  setStatus(message, 'error');
}

function clearDriveError() {
  if (!editorElements.dialogAlert) {
    return;
  }
  editorElements.dialogAlert.hidden = true;
  editorElements.dialogAlert.textContent = '';
}

async function loadGoogleDriveConfig() {
  let config = null;

  try {
    const response = await fetch(driveConfigEndpoint, { cache: 'no-store' });
    if (response.ok) {
      config = await response.json();
    } else if (response.status !== 404) {
      console.warn('Failed to load Google Drive configuration.', response.statusText);
    }
  } catch (error) {
    console.warn('Unable to load Google Drive configuration:', error);
  }

  if (config) {
    if (typeof config.clientId === 'string') {
      googleDriveConfig.clientId = config.clientId.trim();
    }
    if (typeof config.apiKey === 'string') {
      googleDriveConfig.apiKey = config.apiKey.trim();
    }
  }

  if (!googleDriveConfig.clientId) {
    const meta = document.querySelector('meta[name="google-oauth-client-id"]');
    const metaContent = meta?.content?.trim();
    if (metaContent) {
      googleDriveConfig.clientId = metaContent;
    }
  }

  updateDriveConfigMessage();
  updateDriveButtons(Boolean(accessToken));

  if (isDriveConfigured() && gisReady) {
    ensureTokenClient();
  }
}

function isDriveConfigured() {
  return Boolean(googleDriveConfig.clientId && !googleDriveConfig.clientId.startsWith('YOUR_'));
}

function updateDriveConfigMessage() {
  if (!editorElements.driveConfigStatus) {
    return;
  }
  const configured = isDriveConfigured();
  const messageElement = editorElements.driveConfigMessage ?? editorElements.driveConfigStatus;
  const message = configured
    ? 'Google Drive credentials loaded. Sign in to browse Drive files alongside your offline drafts.'
    : 'Provide Google Drive credentials via your secure runtime configuration to enable Drive sync. Offline drafts remain available even without Drive access.';

  editorElements.driveConfigStatus.classList.toggle('configured', configured);
  editorElements.driveConfigStatus.dataset.state = configured ? 'configured' : 'missing';
  messageElement.textContent = message;
}

function clearAccessToken() {
  accessToken = null;
  if (typeof gapi !== 'undefined' && gapi?.client?.setToken) {
    gapi.client.setToken(null);
  }
  if (editorElements.driveFilesBody) {
    editorElements.driveFilesBody.innerHTML = '';
  }
  if (editorElements.driveFilesWrapper) {
    editorElements.driveFilesWrapper.hidden = true;
  }
  driveFolderPath = [{ id: VIRTUAL_ROOT_ID, name: VIRTUAL_ROOT_LABEL, source: FILE_SOURCE_VIRTUAL }];
  currentDriveFolderId = VIRTUAL_ROOT_ID;
  currentFolderSource = FILE_SOURCE_VIRTUAL;
  clearDriveSelection();
  updateDriveButtons(false);
}

function updateDriveButtons(isSignedIn) {
  const hasOfflineFile = currentFileSource === FILE_SOURCE_OFFLINE && Boolean(currentFileId);
  const hasDriveFile = currentFileSource === FILE_SOURCE_DRIVE && Boolean(currentFileId);
  const canQuickSave = hasOfflineFile || (hasDriveFile && isSignedIn);
  if (editorElements.driveOpenButton) {
    editorElements.driveOpenButton.disabled = false;
  }
  if (editorElements.driveSaveButton) {
    editorElements.driveSaveButton.disabled = false;
    editorElements.driveSaveButton.dataset.quickSave = canQuickSave ? 'true' : 'false';
    editorElements.driveSaveButton.setAttribute('aria-disabled', 'false');
  }
  if (editorElements.driveSaveAsButton) {
    editorElements.driveSaveAsButton.disabled = false;
    editorElements.driveSaveAsButton.setAttribute('aria-disabled', 'false');
  }
  if (editorElements.driveRefreshButton) {
    editorElements.driveRefreshButton.disabled = false;
  }
  if (editorElements.driveSignInButton) {
    editorElements.driveSignInButton.hidden = isSignedIn;
  }
  if (editorElements.driveSignOutButton) {
    editorElements.driveSignOutButton.hidden = !isSignedIn;
    editorElements.driveSignOutButton.disabled = !isSignedIn;
    editorElements.driveSignOutButton.setAttribute('aria-disabled', !isSignedIn ? 'true' : 'false');
  }
  if (!editorElements.driveStatus) {
    return;
  }
  if (isSignedIn) {
    editorElements.driveStatus.textContent = 'Connected to Google Drive';
  } else if (!isDriveConfigured()) {
    editorElements.driveStatus.textContent = 'Google Drive credentials not configured';
  } else {
    editorElements.driveStatus.textContent = 'Offline mode – Drive not connected';
  }
}

async function waitForGis() {
  if (gisReady && window.google?.accounts?.oauth2) {
    return;
  }
  await new Promise((resolve) => {
    const check = () => {
      if (gisReady && window.google?.accounts?.oauth2) {
        resolve();
      } else {
        window.setTimeout(check, 100);
      }
    };
    check();
  });
}

function ensureTokenClient() {
  if (tokenClient) {
    return tokenClient;
  }
  if (!window.google?.accounts?.oauth2) {
    return null;
  }
  if (!isDriveConfigured()) {
    return null;
  }
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: googleDriveConfig.clientId,
    scope: scopes,
    callback: () => {}
  });
  return tokenClient;
}

async function requestAccessToken({ forcePrompt = false } = {}) {
  await waitForGis();
  const client = ensureTokenClient();
  if (!client) {
    throw new Error('Google Identity Services is not ready. Please verify your OAuth client ID configuration.');
  }
  return new Promise((resolve, reject) => {
    client.callback = (response) => {
      if (response.error) {
        reject(new Error(response.error_description || response.error));
        return;
      }
      accessToken = response.access_token;
      if (typeof gapi !== 'undefined' && gapi?.client?.setToken) {
        gapi.client.setToken({ access_token: accessToken });
      }
      updateDriveButtons(true);
      resolve(accessToken);
    };
    try {
      const shouldPrompt = forcePrompt || !accessToken;
      client.requestAccessToken({ prompt: shouldPrompt ? 'consent' : '' });
    } catch (error) {
      reject(error);
    }
  });
}

async function ensureDriveAccess({ promptUser = false, forcePrompt = false } = {}) {
  if (!isDriveConfigured()) {
    throw new Error(
      'Provide Google Drive credentials via your secure runtime configuration before connecting to Google Drive.'
    );
  }
  await initializeGapiClient();
  if (!forcePrompt && accessToken) {
    if (typeof gapi !== 'undefined' && gapi?.client?.setToken) {
      gapi.client.setToken({ access_token: accessToken });
    }
    updateDriveButtons(true);
    return accessToken;
  }
  if (!promptUser && !forcePrompt) {
    throw new Error('Sign in to Google Drive to continue.');
  }
  return requestAccessToken({ forcePrompt });
}

async function initializeGapiClient() {
  if (!gapiReady) {
    await waitForGapi();
  }
  if (gapiInitPromise) {
    return gapiInitPromise;
  }
  if (typeof gapi === 'undefined' || !gapi?.client) {
    throw new Error('Google API client library failed to load.');
  }
  gapiInitPromise = gapi.client
    .init({
      discoveryDocs
    })
    .then(() => {
      if (googleDriveConfig.apiKey) {
        gapi.client.setApiKey(googleDriveConfig.apiKey);
      }
    })
    .catch((error) => {
      gapiInitPromise = null;
      throw error;
    });
  await gapiInitPromise;
  return gapiInitPromise;
}

async function waitForGapi() {
  if (gapiReady && window.gapi?.client) {
    return;
  }
  await new Promise((resolve) => {
    const check = () => {
      if (gapiReady && window.gapi?.client) {
        resolve();
      } else {
        window.setTimeout(check, 100);
      }
    };
    check();
  });
}

async function signInToGoogle() {
  clearDriveError();
  try {
    await ensureDriveAccess({ promptUser: true, forcePrompt: true });
    setStatus('Signed in to Google Drive.', 'success');
    refreshDriveFileList({ folderId: currentDriveFolderId, source: currentFolderSource });
  } catch (error) {
    showDriveError(error);
  }
}

function signOutOfGoogle() {
  if (accessToken && window.google?.accounts?.oauth2) {
    try {
      google.accounts.oauth2.revoke(accessToken, () => {});
    } catch (error) {
      console.warn('Failed to revoke Google access token:', error);
    }
  }
  clearDriveError();
  clearAccessToken();
  setStatus('Signed out of Google Drive.');
  refreshDriveFileList({ folderId: currentDriveFolderId, source: currentFolderSource });
}

async function refreshDriveFileList({ folderId = currentDriveFolderId, source = currentFolderSource } = {}) {
  clearDriveError();
  if (!driveFolderPath.length) {
    driveFolderPath = [{ id: VIRTUAL_ROOT_ID, name: VIRTUAL_ROOT_LABEL, source: FILE_SOURCE_VIRTUAL }];
  }
  currentDriveFolderId = folderId || VIRTUAL_ROOT_ID;
  currentFolderSource = source || FILE_SOURCE_VIRTUAL;
  if (editorElements.driveFilesWrapper) {
    editorElements.driveFilesWrapper.hidden = true;
  }
  if (driveDialogMode === 'save') {
    clearDriveSelection();
  }
  if (editorElements.dialogAlert) {
    editorElements.dialogAlert.hidden = true;
    editorElements.dialogAlert.textContent = '';
  }
  let entries = [];
  try {
    if (currentFolderSource === FILE_SOURCE_VIRTUAL) {
      entries = getVirtualRootEntries();
    } else if (currentFolderSource === FILE_SOURCE_OFFLINE) {
      entries = listOfflineEntries(currentDriveFolderId);
    } else if (currentFolderSource === FILE_SOURCE_DRIVE) {
      await ensureDriveAccess({ promptUser: true });
      const query = [
        `'${currentDriveFolderId}' in parents`,
        'trashed=false',
        "(mimeType='application/vnd.google-apps.folder' or mimeType='text/plain' or mimeType='text/markdown' or name contains '.md' or name contains '.markdown')"
      ].join(' and ');
      const response = await gapi.client.drive.files.list({
        pageSize: 100,
        orderBy: 'folder,name_natural',
        q: query,
        fields: 'files(id, name, mimeType, modifiedTime)'
      });
      const files = response.result.files || [];
      entries = files.map((file) => ({ ...file, source: FILE_SOURCE_DRIVE }));
    }
  } catch (error) {
    showDriveError(error);
  }
  populateDriveFiles(entries);
  const hasEntries = entries.length > 0;
  if (editorElements.driveFilesWrapper) {
    editorElements.driveFilesWrapper.hidden = !hasEntries;
  }
  if (!hasEntries && editorElements.dialogAlert) {
    if (currentFolderSource === FILE_SOURCE_OFFLINE) {
      editorElements.dialogAlert.textContent =
        driveDialogMode === 'open'
          ? 'No offline files found yet. Save a document to access it here.'
          : 'This offline folder is empty. Save a file to create it.';
    } else if (currentFolderSource === FILE_SOURCE_DRIVE) {
      editorElements.dialogAlert.textContent =
        driveDialogMode === 'open'
          ? 'No Markdown files found in this Google Drive folder.'
          : 'This Google Drive folder is empty. Save a file here to get started.';
    }
    editorElements.dialogAlert.hidden = !editorElements.dialogAlert.textContent;
  }
  updateDrivePathDisplay();
}

function populateDriveFiles(files) {
  if (!editorElements.driveFilesBody) {
    return;
  }
  editorElements.driveFilesBody.innerHTML = '';
  const isSaveMode = driveDialogMode === 'save';
  files.forEach((file) => {
    const row = document.createElement('tr');
    const nameCell = document.createElement('td');
    const modifiedCell = document.createElement('td');
    const nameWrapper = document.createElement('span');
    nameWrapper.classList.add('entry-name');
    const icon = document.createElement('span');
    icon.classList.add('entry-icon');
    const label = document.createElement('span');
    label.textContent = file.name || 'Untitled';
    nameWrapper.appendChild(icon);
    nameWrapper.appendChild(label);
    nameCell.appendChild(nameWrapper);
    const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
    if (isFolder) {
      row.classList.add('folder-row');
      icon.textContent = '📁';
      modifiedCell.textContent = '';
      row.addEventListener('click', () => {
        enterDriveFolder(file);
      });
    } else {
      icon.textContent = '📄';
      modifiedCell.textContent = formatModifiedTime(file.modifiedTime);
      if (isSaveMode) {
        row.addEventListener('click', () => {
          selectDriveFileForSave(file, row);
        });
      } else {
        row.addEventListener('click', () => {
          if (file.source === FILE_SOURCE_OFFLINE) {
            loadOfflineFile(file.id);
          } else {
            loadDriveFile(file.id, file.name);
          }
          closeDialog();
        });
      }
    }
    row.appendChild(nameCell);
    row.appendChild(modifiedCell);
    editorElements.driveFilesBody.appendChild(row);
  });
}

async function loadDriveFile(fileId, name) {
  clearDriveError();
  try {
    await ensureDriveAccess({ promptUser: true });
    const response = await gapi.client.drive.files.get({ fileId, alt: 'media' });
    const content = response.body || response.result || '';
    currentFileId = fileId;
    const normalizedName = normalizeDisplayName(name);
    currentFileName = normalizedName;
    pendingFileName = normalizedName;
    currentFileSource = FILE_SOURCE_DRIVE;
    currentOfflineParentId = OFFLINE_ROOT_ID;
    updateTitleInput();
    applyEditorUpdate(content, content.length, content.length, { markDirty: false, resetHistory: true });
    localStorage.setItem(
      'markdown-editor-current-file',
      JSON.stringify({ id: fileId, name: normalizedName, source: FILE_SOURCE_DRIVE })
    );
    setStatus(`Loaded ${normalizedName} from Google Drive.`, 'success');
    updateDriveButtons(true);
  } catch (error) {
    showDriveError(error);
  }
}

function loadOfflineFile(fileId) {
  clearDriveError();
  const file = readOfflineFile(fileId);
  if (!file) {
    showDriveError('Unable to load the selected offline file.');
    return;
  }
  const content = file.content ?? '';
  currentFileId = file.id;
  currentFileSource = FILE_SOURCE_OFFLINE;
  currentOfflineParentId = file.parentId || OFFLINE_ROOT_ID;
  const normalizedName = normalizeDisplayName(file.name);
  currentFileName = normalizedName;
  pendingFileName = normalizedName;
  updateTitleInput();
  applyEditorUpdate(content, content.length, content.length, { markDirty: false, resetHistory: true });
  localStorage.setItem(
    'markdown-editor-current-file',
    JSON.stringify({
      id: file.id,
      name: normalizedName,
      source: FILE_SOURCE_OFFLINE,
      parentId: currentOfflineParentId
    })
  );
  setStatus(`Loaded ${normalizedName} from offline storage.`, 'success');
  updateDriveButtons(Boolean(accessToken));
}

function prepareContentForSaving() {
  let content = markdownContent;
  if (editorMode === 'html') {
    const latestHtml = getHtmlEditorContent();
    htmlContent = latestHtml;
    const converted = convertHtmlToMarkdown(latestHtml);
    markdownContent = converted;
    updateCounts(converted);
    content = converted;
  }
  persistEditorContent(content, { immediate: true });
  flushPendingContentPersistence();
  return content;
}

function saveToOffline({ fileId = null, fileName = null, folderId = null } = {}) {
  const content = prepareContentForSaving();
  const targetFolderId = folderId || currentOfflineParentId || OFFLINE_ROOT_ID;
  const targetName = ensureMarkdownExtension(fileName ?? getPendingFileNameForSaving());
  const saved = writeOfflineFile({ fileId, fileName: targetName, content, folderId: targetFolderId });
  if (!saved) {
    return null;
  }
  currentFileId = saved.id;
  const normalizedName = normalizeDisplayName(saved.name);
  currentFileName = normalizedName;
  pendingFileName = normalizedName;
  currentFileSource = FILE_SOURCE_OFFLINE;
  currentOfflineParentId = saved.parentId || OFFLINE_ROOT_ID;
  isDirty = false;
  updateTitleInput();
  updateFileIndicator();
  localStorage.setItem(
    'markdown-editor-current-file',
    JSON.stringify({
      id: currentFileId,
      name: currentFileName,
      source: FILE_SOURCE_OFFLINE,
      parentId: currentOfflineParentId
    })
  );
  setStatus(`Saved ${currentFileName} offline.`, 'success');
  updateDriveButtons(Boolean(accessToken));
  return saved;
}

async function saveToDrive({ forceNew = false, folderId = null, fileName = null, fileId = null } = {}) {
  clearDriveError();
  const content = prepareContentForSaving();
  try {
    await ensureDriveAccess({ promptUser: true });
    let targetFileId = fileId ?? (forceNew ? null : currentFileId);
    let targetFileName = fileName ?? getPendingFileNameForSaving();
    if ((forceNew && !fileName) || (!targetFileId && !fileName)) {
      const suggestedName = ensureMarkdownExtension(targetFileName);
      const input = window.prompt('File name', suggestedName);
      if (!input) {
        return false;
      }
      targetFileName = ensureMarkdownExtension(input);
    } else {
      targetFileName = ensureMarkdownExtension(targetFileName);
    }
    const result = await uploadToDrive(targetFileId, targetFileName, content, targetFileId ? null : folderId);
    currentFileId = result.id;
    const savedName = normalizeDisplayName(result?.name || targetFileName);
    currentFileName = savedName;
    pendingFileName = savedName;
    currentFileSource = FILE_SOURCE_DRIVE;
    currentOfflineParentId = OFFLINE_ROOT_ID;
    isDirty = false;
    updateTitleInput();
    updateFileIndicator();
    localStorage.setItem(
      'markdown-editor-current-file',
      JSON.stringify({ id: currentFileId, name: savedName, source: FILE_SOURCE_DRIVE })
    );
    setStatus(`Saved ${savedName} to Google Drive.`, 'success');
    updateDriveButtons(true);
    return true;
  } catch (error) {
    showDriveError(error);
    return false;
  }
}

async function handleSaveButtonClick() {
  if (currentFileSource === FILE_SOURCE_OFFLINE && currentFileId) {
    saveToOffline({ fileId: currentFileId, folderId: currentOfflineParentId, fileName: currentFileName });
    return;
  }
  if (currentFileSource === FILE_SOURCE_DRIVE && currentFileId) {
    await saveToDrive({ fileId: currentFileId, fileName: currentFileName });
    return;
  }
  openDialog('save');
}

async function uploadToDrive(fileId, fileName, content, folderId = null) {
  const boundary = '-------314159265358979323846';
  const delimiter = `\r\n--${boundary}\r\n`;
  const closeDelimiter = `\r\n--${boundary}--`;
  const metadata = {
    name: fileName,
    mimeType: 'text/plain'
  };
  if (!fileId && folderId) {
    metadata.parents = [folderId];
  }
  const multipartRequestBody =
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    'Content-Type: text/plain\r\n\r\n' +
    content +
    closeDelimiter;

  const path = fileId ? `/upload/drive/v3/files/${fileId}` : '/upload/drive/v3/files';
  const method = fileId ? 'PATCH' : 'POST';
  const request = {
    path,
    method,
    params: { uploadType: 'multipart', fields: 'id,name' },
    headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
    body: multipartRequestBody
  };

  const response = await gapi.client.request(request);
  return response.result;
}

function restoreLastFile() {
  const stored = JSON.parse(localStorage.getItem('markdown-editor-current-file') || 'null');
  if (stored?.name) {
    currentFileName = normalizeDisplayName(stored.name);
    currentFileId = stored.id || null;
    currentFileSource = stored.source || (stored.id ? FILE_SOURCE_DRIVE : null);
    if (currentFileSource === FILE_SOURCE_OFFLINE) {
      currentOfflineParentId = stored.parentId || OFFLINE_ROOT_ID;
    } else {
      currentOfflineParentId = OFFLINE_ROOT_ID;
      if (currentFileSource !== FILE_SOURCE_DRIVE && currentFileId) {
        currentFileSource = FILE_SOURCE_DRIVE;
      }
    }
  } else {
    currentFileName = normalizeDisplayName(currentFileName);
    currentFileId = null;
    currentFileSource = null;
    currentOfflineParentId = OFFLINE_ROOT_ID;
  }
  pendingFileName = currentFileName;
  updateTitleInput();
  updateFileIndicator();
  updateDriveButtons(Boolean(accessToken));
}

window.onGapiLoaded = () => {
  if (typeof gapi === 'undefined' || !gapi?.load) {
    console.error('Google API platform script loaded without gapi.load available.');
    return;
  }

  gapi.load('client', {
    callback: () => {
      gapiReady = true;
    },
    onerror: () => {
      console.error('Failed to load Google API client library modules.');
    }
  });
};

window.onGoogleAccountsLoaded = () => {
  gisReady = true;
  if (isDriveConfigured()) {
    ensureTokenClient();
  }
  updateDriveConfigMessage();
};

window.addEventListener('beforeunload', () => {
  flushPendingContentPersistence();
});

window.addEventListener('pagehide', () => {
  flushPendingContentPersistence();
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') {
    flushPendingContentPersistence();
  }
});

document.addEventListener('keydown', (event) => {
  const editor = editorElements.editor;
  const isModifier = event.ctrlKey || event.metaKey;
  if (!editor || !isModifier) {
    return;
  }
  const target = event.target;
  if (target instanceof Node && !editor.contains(target)) {
    return;
  }
  const key = event.key.toLowerCase();
  if (['b', 'i', 'y', 'z'].includes(key) && editorMode !== 'markdown') {
    return;
  }
  switch (key) {
    case 'b':
      event.preventDefault();
      applyMarkdown('bold');
      break;
    case 'i':
      event.preventDefault();
      applyMarkdown('italic');
      break;
    case 'y':
      event.preventDefault();
      redo();
      break;
    case 'z':
      event.preventDefault();
      if (event.shiftKey) {
        redo();
      } else {
        undo();
      }
      break;
    default:
      break;
  }
});

initializeThemePreference();

init().catch((error) => {
  console.error('Failed to initialise the Markdown editor.', error);
  setStatus('Failed to initialise Google Drive integration. Check console for details.', 'error');
});
