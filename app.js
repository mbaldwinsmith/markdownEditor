'use strict';

const editorElements = {
  editor: document.getElementById('markdown-input'),
  wordCount: document.getElementById('word-count'),
  charCount: document.getElementById('char-count'),
  fileIndicator: document.getElementById('current-file'),
  statusMessage: document.getElementById('status-message'),
  driveStatus: document.getElementById('drive-status'),
  toolbarButtons: document.querySelectorAll('[data-action]'),
  dialog: document.getElementById('drive-dialog'),
  dialogClose: document.getElementById('drive-dialog-close'),
  dialogCancel: document.getElementById('drive-dialog-cancel'),
  dialogAlert: document.getElementById('drive-alert'),
  driveFilesWrapper: document.getElementById('drive-files'),
  driveFilesBody: document.getElementById('drive-files-body'),
  driveRefreshButton: document.getElementById('drive-refresh-files'),
  driveOpenButton: document.getElementById('drive-open'),
  driveSaveButton: document.getElementById('drive-save'),
  driveSaveAsButton: document.getElementById('drive-save-as'),
  driveSignInButton: document.getElementById('drive-sign-in'),
  driveSignOutButton: document.getElementById('drive-sign-out'),
  driveConfigStatus: document.getElementById('drive-config-status')
};

let currentFileId = null;
let currentFileName = 'Untitled.md';
let isDirty = true;
let gapiReady = false;
let gapiInitPromise = null;
let editorContent = '';
let lastSelection = { start: 0, end: 0 };
let tokenClient = null;
let gisReady = false;
let accessToken = null;

const discoveryDocs = ['https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'];
const scopes = 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/drive.file';

const googleDriveConfig = Object.freeze({
  clientId: document.querySelector('meta[name="google-oauth-client-id"]')?.content?.trim() ?? '',
  apiKey: document.querySelector('meta[name="google-api-key"]')?.content?.trim() ?? ''
});

const defaultMarkdown = `# Welcome to the Markdown Editor PWA

Start typing in the editor to craft your Markdown documents. Use the toolbar buttons to quickly insert Markdown formatting such as **bold**, *italic*, links, lists, tables, and more.

## Features

- Minimal editor focused on Markdown syntax
- Word and character counts update automatically
- Save your documents to Google Drive
- Install the app to work offline as a Progressive Web App

> Tip: Update the \`google-oauth-client-id\` meta tag in `index.html` with your OAuth client ID to enable Google Drive sync.
`;

function init() {
  const savedContent = localStorage.getItem('markdown-editor-content');
  const initialContent = savedContent ?? defaultMarkdown;
  applyEditorUpdate(initialContent, initialContent.length, initialContent.length, {
    persistContent: false,
    markDirty: false,
    focus: false
  });
  restoreLastFile();
  updateDriveButtons(false);
  updateDriveConfigMessage();
  if (!isDriveConfigured()) {
    setStatus('Configure your Google OAuth client ID in index.html to enable Google Drive sync.', 'error');
  }
  attachEventListeners();
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

  const fragment = document.createDocumentFragment();
  const lines = content.length ? content.split(/\n/u) : [''];

  lines.forEach((line, index) => {
    const lineElement = document.createElement('div');
    lineElement.classList.add('editor-line');

    const headingMatch = line.match(/^(#{1,6})\s+(.*)$/u);
    if (headingMatch) {
      const level = Math.min(headingMatch[1].length, 6);
      lineElement.classList.add(`heading-${level}`);
    }

    if (!line) {
      lineElement.innerHTML = '&#8203;';
    } else {
      lineElement.textContent = '';
      lineElement.appendChild(renderInlineMarkdown(line));
    }

    fragment.appendChild(lineElement);

    if (index < lines.length - 1) {
      fragment.appendChild(document.createTextNode('\n'));
    }
  });

  editor.innerHTML = '';
  editor.appendChild(fragment);
}

function getPlainTextFromEditor() {
  const editor = editorElements.editor;
  if (!editor) {
    return '';
  }
  return (editor.textContent || '').replace(/\u200B/gu, '');
}

function getSelectionOffsets() {
  const editor = editorElements.editor;
  if (!editor) {
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

  const preRange = range.cloneRange();
  preRange.selectNodeContents(editor);
  preRange.setEnd(range.startContainer, range.startOffset);
  const start = preRange.toString().replace(/\u200B/gu, '').length;
  const selectedLength = range.toString().replace(/\u200B/gu, '').length;

  return { start, end: start + selectedLength };
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

function setSelectionRange(start, end) {
  const editor = editorElements.editor;
  if (!editor) {
    return;
  }

  const totalLength = editorContent.length;
  const clampedStart = Math.max(0, Math.min(start, totalLength));
  const clampedEnd = Math.max(0, Math.min(end, totalLength));

  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const range = document.createRange();
  const startPosition = resolveOffset(clampedStart);
  const endPosition = resolveOffset(clampedEnd);

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

  editorContent = content;
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
  const { start, end } = getSelectionOffsets();
  const value = getPlainTextFromEditor();
  applyEditorUpdate(value, start, end, { focus: false });
}

function updateSelectionCache() {
  const editor = editorElements.editor;
  if (!editor) {
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
  const indicator = `${currentFileName}${isDirty ? ' • Unsaved changes' : ''}`;
  editorElements.fileIndicator.textContent = indicator;
}

function attachEventListeners() {
  const editor = editorElements.editor;
  editor.addEventListener('input', () => handleEditorInput());
  editor.addEventListener('keyup', () => updateSelectionCache());
  editor.addEventListener('mouseup', () => updateSelectionCache());
  editor.addEventListener('blur', () => updateSelectionCache());

  document.addEventListener('selectionchange', () => updateSelectionCache());

  editorElements.toolbarButtons.forEach((button) => {
    button.addEventListener('mousedown', (event) => event.preventDefault());
    button.addEventListener('click', () => applyMarkdown(button.dataset.action));
  });

  editorElements.dialogClose.addEventListener('click', () => closeDialog());
  editorElements.dialogCancel.addEventListener('click', () => closeDialog());

  if (editorElements.driveRefreshButton) {
    editorElements.driveRefreshButton.addEventListener('click', () => {
      refreshDriveFileList();
    });
  }

  editorElements.driveOpenButton.addEventListener('click', () => {
    openDialog();
    refreshDriveFileList();
  });

  editorElements.driveSaveButton.addEventListener('click', () => {
    saveToDrive();
  });

  editorElements.driveSaveAsButton.addEventListener('click', () => {
    saveToDrive(true);
  });

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
  const value = editorContent;
  const selected = value.slice(start, end) || placeholder;
  const newValue = `${value.slice(0, start)}${before}${selected}${after}${value.slice(end)}`;
  const newStart = start + before.length;
  const newEnd = newStart + selected.length;
  applyEditorUpdate(newValue, newStart, newEnd);
}

function insertSnippet(snippet) {
  const { start, end } = getSelectionOffsets();
  const value = editorContent;
  const newValue = `${value.slice(0, start)}${snippet}${value.slice(end)}`;
  const cursorPosition = start + snippet.length;
  applyEditorUpdate(newValue, cursorPosition, cursorPosition);
}

function applyLinePrefix(prefix, placeholder = '') {
  const value = editorContent;
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
  const value = editorContent;
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

function openDialog() {
  clearDriveError();
  updateDriveConfigMessage();
  editorElements.dialog.classList.add('active');
  editorElements.dialog.setAttribute('aria-hidden', 'false');
  editorElements.driveFilesWrapper.hidden = true;
  editorElements.driveRefreshButton?.focus();
}

function closeDialog() {
  editorElements.dialog.classList.remove('active');
  editorElements.dialog.setAttribute('aria-hidden', 'true');
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

function isDriveConfigured() {
  return Boolean(googleDriveConfig.clientId && !googleDriveConfig.clientId.startsWith('YOUR_'));
}

function updateDriveConfigMessage() {
  if (!editorElements.driveConfigStatus) {
    return;
  }
  if (isDriveConfigured()) {
    editorElements.driveConfigStatus.textContent =
      'Your OAuth client ID is configured. Sign in to browse Google Drive files.';
  } else {
    editorElements.driveConfigStatus.textContent =
      'Set the google-oauth-client-id meta tag in index.html to your OAuth client ID to enable Drive sync.';
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
    editorElements.driveStatus.textContent = 'OAuth client ID not configured';
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
    throw new Error('Set the google-oauth-client-id meta tag in index.html before connecting to Google Drive.');
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

async function refreshDriveFileList() {
  clearDriveError();
  if (editorElements.driveFilesWrapper) {
    editorElements.driveFilesWrapper.hidden = true;
  }
  try {
    await ensureDriveAccess({ promptUser: true });
    const response = await gapi.client.drive.files.list({
      pageSize: 50,
      orderBy: 'modifiedTime desc',
      q: "mimeType='text/plain' or name contains '.md'",
      fields: 'files(id, name, modifiedTime)'
    });
    const files = response.result.files || [];
    populateDriveFiles(files);
    editorElements.driveFilesWrapper.hidden = files.length === 0;
    if (files.length === 0) {
      editorElements.driveFilesWrapper.hidden = true;
      editorElements.dialogAlert.textContent = 'No compatible files found in Google Drive.';
      editorElements.dialogAlert.hidden = false;
    }
  } catch (error) {
    showDriveError(error);
  }
}

function populateDriveFiles(files) {
  editorElements.driveFilesBody.innerHTML = '';
  files.forEach((file) => {
    const row = document.createElement('tr');
    const nameCell = document.createElement('td');
    const modifiedCell = document.createElement('td');
    nameCell.textContent = file.name;
    modifiedCell.textContent = file.modifiedTime ? new Date(file.modifiedTime).toLocaleString() : '';
    row.appendChild(nameCell);
    row.appendChild(modifiedCell);
    row.addEventListener('click', () => {
      loadDriveFile(file.id, file.name);
      closeDialog();
    });
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
    currentFileName = name;
    applyEditorUpdate(content, content.length, content.length, { markDirty: false });
    localStorage.setItem('markdown-editor-current-file', JSON.stringify({ id: fileId, name }));
    setStatus(`Loaded ${name} from Google Drive.`, 'success');
  } catch (error) {
    showDriveError(error);
  }
}

async function saveToDrive(forceNew = false) {
  clearDriveError();
  try {
    await ensureDriveAccess({ promptUser: true });
    let fileId = currentFileId;
    let fileName = currentFileName;
    if (forceNew || !fileId) {
      const suggestedName = currentFileName || 'Untitled.md';
      const input = window.prompt('File name', suggestedName);
      if (!input) {
        return;
      }
      fileName = input.endsWith('.md') ? input : `${input}.md`;
      fileId = forceNew ? null : currentFileId;
    }
    const content = editorContent;
    const result = await uploadToDrive(fileId, fileName, content);
    currentFileId = result.id;
    currentFileName = result.name;
    isDirty = false;
    updateFileIndicator();
    localStorage.setItem('markdown-editor-current-file', JSON.stringify({ id: currentFileId, name: currentFileName }));
    setStatus(`Saved ${currentFileName} to Google Drive.`, 'success');
  } catch (error) {
    showDriveError(error);
  }
}

async function uploadToDrive(fileId, fileName, content) {
  const boundary = '-------314159265358979323846';
  const delimiter = `\r\n--${boundary}\r\n`;
  const closeDelimiter = `\r\n--${boundary}--`;
  const metadata = {
    name: fileName,
    mimeType: 'text/plain'
  };
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
    currentFileName = stored.name;
    currentFileId = stored.id || null;
  }
  updateFileIndicator();
}

window.onGapiLoaded = () => {
  gapiReady = true;
};

window.onGoogleAccountsLoaded = () => {
  gisReady = true;
  if (isDriveConfigured()) {
    ensureTokenClient();
  }
  updateDriveConfigMessage();
};

document.addEventListener('keydown', (event) => {
  const isModifier = event.ctrlKey || event.metaKey;
  if (event.target === editorElements.editor && isModifier) {
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
  }
});

init();
