'use strict';

const editorElements = {
  editor: document.getElementById('markdown-input'),
  fileTitleInput: document.getElementById('file-title-input'),
  wordCount: document.getElementById('word-count'),
  charCount: document.getElementById('char-count'),
  fileIndicator: document.getElementById('current-file'),
  statusMessage: document.getElementById('status-message'),
  driveStatus: document.getElementById('drive-status'),
  tocList: document.getElementById('toc-list'),
  tocEmptyState: document.getElementById('toc-empty'),
  toolbarButtons: document.querySelectorAll('[data-action]'),
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
  driveConfigStatus: document.getElementById('drive-config-status')
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

let driveDialogMode = 'open';
let currentDriveFolderId = DRIVE_ROOT_ID;
let driveFolderPath = [];
let pendingSaveFileId = null;
let pendingSaveFileName = '';

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
  const savedContent = localStorage.getItem('markdown-editor-content');
  const initialContent = savedContent ?? defaultMarkdown;
  applyEditorUpdate(initialContent, initialContent.length, initialContent.length, {
    persistContent: false,
    markDirty: false,
    focus: false
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
  const pattern = /\*\*([^*]+)\*\*|\*([^*]+)\*/gu;
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
    tocContainer.hidden = false;
  }
  const fragment = document.createDocumentFragment();
  const lines = content.length ? content.split(/\n/u) : [''];
  const headings = [];
  const slugCounts = new Map();

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
  updateTableOfContents(headings);
}

function renderHtmlEditor(content) {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  editor.classList.add('html-mode');
  editor.textContent = content || '';
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

function applyEditorUpdate(content, selectionStart = content.length, selectionEnd = selectionStart, options = {}) {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  const { markDirty = true, persistContent = true, focus = true } = options;
  const previousScrollTop = editor.scrollTop;
  const previousScrollLeft = editor.scrollLeft;

  editorMode = 'markdown';
  markdownContent = content;
  renderFormattedMarkdown(content);

  editor.scrollTop = previousScrollTop;
  editor.scrollLeft = previousScrollLeft;

  updateCounts(content);

  if (persistContent) {
    localStorage.setItem('markdown-editor-content', content);
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
}

function handleEditorInput() {
  if (editorMode === 'html') {
    const htmlValue = getHtmlEditorContent();
    htmlContent = htmlValue;
    const markdownValue = convertHtmlToMarkdown(htmlValue);
    markdownContent = markdownValue;
    updateCounts(markdownValue);
    localStorage.setItem('markdown-editor-content', markdownValue);
    isDirty = true;
    updateFileIndicator();
    return;
  }

  const { start, end } = getSelectionOffsets();
  const value = getPlainTextFromEditor();
  applyEditorUpdate(value, start, end, { focus: false });
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
  const indicator = `${displayName}${isDirty || hasTitleChanges() ? ' • Unsaved changes' : ''}`;
  editorElements.fileIndicator.textContent = indicator;
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
    saveToDrive();
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
    case 'bold':
      wrapSelection('**', '**', 'bold text');
      break;
    case 'italic':
      wrapSelection('*', '*', 'italic text');
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
  driveFolderPath = [{ id: DRIVE_ROOT_ID, name: DRIVE_ROOT_LABEL }];
  currentDriveFolderId = DRIVE_ROOT_ID;
  pendingSaveFileId = null;
  pendingSaveFileName = '';
  updateDriveDialogMode();
  refreshDriveFileList({ folderId: currentDriveFolderId });
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
}

function updateDriveDialogMode() {
  const isSaveMode = driveDialogMode === 'save';
  if (editorElements.driveDialogTitle) {
    editorElements.driveDialogTitle.textContent = isSaveMode
      ? 'Save to Google Drive'
      : 'Open from Google Drive';
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
    driveFolderPath = [{ id: DRIVE_ROOT_ID, name: DRIVE_ROOT_LABEL }];
  }
  editorElements.driveBreadcrumbs.innerHTML = '';
  driveFolderPath.forEach((entry, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = entry.name || 'Untitled';
    button.disabled = index === driveFolderPath.length - 1;
    if (!button.disabled) {
      button.addEventListener('click', () => {
        driveFolderPath = driveFolderPath.slice(0, index + 1);
        const target = driveFolderPath[driveFolderPath.length - 1];
        currentDriveFolderId = target.id;
        clearDriveSelection();
        refreshDriveFileList({ folderId: target.id });
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
  clearDriveSelection();
  refreshDriveFileList({ folderId: parent.id });
}

function enterDriveFolder(folder) {
  if (!folder?.id) {
    return;
  }
  const name = folder.name || 'Untitled folder';
  driveFolderPath = [...driveFolderPath, { id: folder.id, name }];
  currentDriveFolderId = folder.id;
  clearDriveSelection();
  refreshDriveFileList({ folderId: folder.id });
}

function clearDriveSelection() {
  pendingSaveFileId = null;
  pendingSaveFileName = '';
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
  const options = {
    folderId: currentDriveFolderId,
    fileName
  };
  if (pendingSaveFileId) {
    options.fileId = pendingSaveFileId;
  } else {
    options.forceNew = true;
  }
  const confirmButton = editorElements.driveSaveConfirmButton;
  if (confirmButton) {
    confirmButton.disabled = true;
  }
  const success = await saveToDrive(options);
  if (confirmButton) {
    confirmButton.disabled = false;
  }
  if (success) {
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
  editorElements.dialogAlert.textContent = message;
  editorElements.dialogAlert.hidden = false;
  setStatus(message, 'error');
}

function clearDriveError() {
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
  if (isDriveConfigured()) {
    editorElements.driveConfigStatus.textContent =
      'Google Drive credentials loaded. Sign in to browse your files.';
  } else {
    editorElements.driveConfigStatus.textContent =
      'Provide Google Drive credentials via your secure runtime configuration to enable Drive sync. For local development you may set the google-oauth-client-id meta tag.';
  }
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
  updateDriveButtons(false);
}

function updateDriveButtons(isSignedIn) {
  const disabled = !isSignedIn;
  editorElements.driveOpenButton.disabled = disabled;
  editorElements.driveSaveButton.disabled = disabled;
  editorElements.driveSaveAsButton.disabled = disabled;
  if (editorElements.driveRefreshButton) {
    editorElements.driveRefreshButton.disabled = disabled;
  }
  editorElements.driveSignInButton.hidden = isSignedIn;
  editorElements.driveSignOutButton.hidden = !isSignedIn;
  if (!isSignedIn && !isDriveConfigured()) {
    editorElements.driveStatus.textContent = 'Google Drive credentials not configured';
  } else {
    editorElements.driveStatus.textContent = isSignedIn ? 'Connected to Google Drive' : 'Not connected';
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
}

async function refreshDriveFileList({ folderId = currentDriveFolderId } = {}) {
  clearDriveError();
  if (!driveFolderPath.length) {
    driveFolderPath = [{ id: DRIVE_ROOT_ID, name: DRIVE_ROOT_LABEL }];
  }
  currentDriveFolderId = folderId || DRIVE_ROOT_ID;
  if (editorElements.driveFilesWrapper) {
    editorElements.driveFilesWrapper.hidden = true;
  }
  if (driveDialogMode === 'save') {
    clearDriveSelection();
  }
  try {
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
    populateDriveFiles(files);
    const hasEntries = files.length > 0;
    if (editorElements.driveFilesWrapper) {
      editorElements.driveFilesWrapper.hidden = !hasEntries;
    }
    if (!hasEntries) {
      editorElements.dialogAlert.textContent =
        driveDialogMode === 'open'
          ? 'No Markdown files found in this folder.'
          : 'This folder is empty. Save a file here to get started.';
      editorElements.dialogAlert.hidden = false;
    }
    updateDrivePathDisplay();
  } catch (error) {
    showDriveError(error);
  }
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
      modifiedCell.textContent = file.modifiedTime ? new Date(file.modifiedTime).toLocaleString() : '';
      if (isSaveMode) {
        row.addEventListener('click', () => {
          selectDriveFileForSave(file, row);
        });
      } else {
        row.addEventListener('click', () => {
          loadDriveFile(file.id, file.name);
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
    updateTitleInput();
    applyEditorUpdate(content, content.length, content.length, { markDirty: false });
    localStorage.setItem('markdown-editor-current-file', JSON.stringify({ id: fileId, name: normalizedName }));
    setStatus(`Loaded ${normalizedName} from Google Drive.`, 'success');
  } catch (error) {
    showDriveError(error);
  }
}

async function saveToDrive({ forceNew = false, folderId = null, fileName = null, fileId = null } = {}) {
  clearDriveError();
  if (editorMode === 'html') {
    const latestHtml = getHtmlEditorContent();
    htmlContent = latestHtml;
    const converted = convertHtmlToMarkdown(latestHtml);
    markdownContent = converted;
    updateCounts(converted);
    localStorage.setItem('markdown-editor-content', converted);
  }
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
    const content = markdownContent;
    const result = await uploadToDrive(targetFileId, targetFileName, content, targetFileId ? null : folderId);
    currentFileId = result.id;
    const savedName = normalizeDisplayName(result?.name || targetFileName);
    currentFileName = savedName;
    pendingFileName = savedName;
    isDirty = false;
    updateTitleInput();
    updateFileIndicator();
    localStorage.setItem('markdown-editor-current-file', JSON.stringify({ id: currentFileId, name: savedName }));
    setStatus(`Saved ${savedName} to Google Drive.`, 'success');
    return true;
  } catch (error) {
    showDriveError(error);
    return false;
  }
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
  } else {
    currentFileName = normalizeDisplayName(currentFileName);
  }
  pendingFileName = currentFileName;
  updateTitleInput();
  updateFileIndicator();
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
  switch (event.key.toLowerCase()) {
    case 'b':
      event.preventDefault();
      applyMarkdown('bold');
      break;
    case 'i':
      event.preventDefault();
      applyMarkdown('italic');
      break;
    default:
      break;
  }
});

init().catch((error) => {
  console.error('Failed to initialise the Markdown editor.', error);
  setStatus('Failed to initialise Google Drive integration. Check console for details.', 'error');
});
