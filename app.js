'use strict';

const editorElements = {
  textarea: document.getElementById('markdown-input'),
  preview: document.getElementById('preview'),
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
  driveSettingsButton: document.getElementById('drive-settings'),
  driveConnectButton: document.getElementById('drive-connect'),
  driveRefreshButton: document.getElementById('drive-refresh-files'),
  driveOpenButton: document.getElementById('drive-open'),
  driveSaveButton: document.getElementById('drive-save'),
  driveSaveAsButton: document.getElementById('drive-save-as'),
  driveSignInButton: document.getElementById('drive-sign-in'),
  driveSignOutButton: document.getElementById('drive-sign-out'),
  apiKeyInput: document.getElementById('drive-api-key'),
  clientIdInput: document.getElementById('drive-client-id')
};

let currentFileId = null;
let currentFileName = 'Untitled.md';
let isDirty = true;
let gapiReady = false;
let gapiInitPromise = null;
let googleAuthInstance = null;

const discoveryDocs = ['https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'];
const scopes = 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/drive.file';

const defaultMarkdown = `# Welcome to the Markdown Editor PWA

Start typing in the editor on the left to see the rendered Markdown preview on the right. Use the toolbar buttons to quickly insert Markdown formatting such as **bold**, *italic*, links, lists, tables, and more.

## Features

- Live preview rendered with [Marked](https://marked.js.org/)
- Word and character counts update automatically
- Save your documents to Google Drive
- Install the app to work offline as a Progressive Web App

> Tip: provide your Google API key and OAuth client ID in the Drive settings dialog to enable cloud sync.
`;

function init() {
  if (typeof marked !== 'undefined') {
    marked.setOptions({
      breaks: true,
      gfm: true,
      headerIds: true
    });
  }

  const savedContent = localStorage.getItem('markdown-editor-content');
  editorElements.textarea.value = savedContent ?? defaultMarkdown;
  renderPreview(editorElements.textarea.value);
  updateCounts(editorElements.textarea.value);
  loadStoredCredentials();
  restoreLastFile();
  updateDriveButtons(false);
  attachEventListeners();
  registerServiceWorker();
}

function renderPreview(markdownText) {
  if (typeof marked === 'undefined' || typeof DOMPurify === 'undefined') {
    editorElements.preview.textContent = markdownText;
    return;
  }

  try {
    const html = marked.parse(markdownText);
    editorElements.preview.innerHTML = DOMPurify.sanitize(html, { USE_PROFILES: { html: true } });
  } catch (error) {
    editorElements.preview.textContent = markdownText;
    console.error('Failed to render markdown', error);
  }
}

function updateCounts(content) {
  const words = content.trim() ? content.trim().split(/\s+/u).length : 0;
  const characters = content.length;
  editorElements.wordCount.textContent = words;
  editorElements.charCount.textContent = characters;
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
  editorElements.textarea.addEventListener('input', () => {
    const value = editorElements.textarea.value;
    renderPreview(value);
    updateCounts(value);
    isDirty = true;
    localStorage.setItem('markdown-editor-content', value);
    updateFileIndicator();
  });

  editorElements.toolbarButtons.forEach((button) => {
    button.addEventListener('click', () => applyMarkdown(button.dataset.action));
  });

  editorElements.driveSettingsButton.addEventListener('click', () => openDialog());
  editorElements.dialogClose.addEventListener('click', () => closeDialog());
  editorElements.dialogCancel.addEventListener('click', () => closeDialog());

  editorElements.driveConnectButton.addEventListener('click', () => {
    saveCredentials();
    initializeGapiClient().catch((error) => showDriveError(error));
  });

  editorElements.driveRefreshButton.addEventListener('click', () => {
    refreshDriveFileList();
  });

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
  const textarea = editorElements.textarea;
  textarea.focus();

  switch (action) {
    case 'bold':
      wrapSelection('**', '**', 'bold text');
      break;
    case 'italic':
      wrapSelection('*', '*', 'italic text');
      break;
    case 'heading':
      applyLinePrefix('## ');
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
  const textarea = editorElements.textarea;
  const { selectionStart, selectionEnd, value } = textarea;
  const selected = value.slice(selectionStart, selectionEnd) || placeholder;
  const newText = `${before}${selected}${after}`;
  textarea.setRangeText(newText, selectionStart, selectionEnd, 'end');

  const newStart = selectionStart + before.length;
  const newEnd = newStart + selected.length;
  textarea.setSelectionRange(newStart, newEnd);
  triggerEditorUpdate();
}

function insertSnippet(snippet) {
  const textarea = editorElements.textarea;
  const { selectionStart, selectionEnd } = textarea;
  textarea.setRangeText(snippet, selectionStart, selectionEnd, 'end');
  const cursorPosition = selectionStart + snippet.length;
  textarea.setSelectionRange(cursorPosition, cursorPosition);
  triggerEditorUpdate();
}

function applyLinePrefix(prefix) {
  const textarea = editorElements.textarea;
  const { selectionStart, selectionEnd, value } = textarea;
  const lineStart = value.lastIndexOf('\n', selectionStart - 1) + 1;
  let lineEnd = value.indexOf('\n', selectionEnd);
  if (lineEnd === -1) {
    lineEnd = value.length;
  }
  const selected = value.slice(lineStart, lineEnd);
  const lines = selected.split('\n');
  const updatedLines = lines.map((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      return prefix;
    }
    if (line.startsWith(prefix)) {
      return line;
    }
    return `${prefix}${line.replace(/^#{1,6}\s+/u, '')}`;
  });
  const updated = updatedLines.join('\n');
  textarea.setRangeText(updated, lineStart, lineEnd, 'end');
  textarea.setSelectionRange(lineStart, lineStart + updated.length);
  triggerEditorUpdate();
}

function applyList(ordered) {
  const textarea = editorElements.textarea;
  const { selectionStart, selectionEnd, value } = textarea;
  const lineStart = value.lastIndexOf('\n', selectionStart - 1) + 1;
  let lineEnd = value.indexOf('\n', selectionEnd);
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
  textarea.setRangeText(updated, lineStart, lineEnd, 'end');
  textarea.setSelectionRange(lineStart, lineStart + updated.length);
  triggerEditorUpdate();
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

function triggerEditorUpdate() {
  const value = editorElements.textarea.value;
  renderPreview(value);
  updateCounts(value);
  isDirty = true;
  localStorage.setItem('markdown-editor-content', value);
  updateFileIndicator();
}

function openDialog() {
  clearDriveError();
  editorElements.dialog.classList.add('active');
  editorElements.dialog.setAttribute('aria-hidden', 'false');
  editorElements.driveFilesWrapper.hidden = true;
  editorElements.dialog.querySelector('input')?.focus();
}

function closeDialog() {
  editorElements.dialog.classList.remove('active');
  editorElements.dialog.setAttribute('aria-hidden', 'true');
}

function loadStoredCredentials() {
  const stored = JSON.parse(localStorage.getItem('markdown-editor-drive-credentials') || '{}');
  if (stored.apiKey) {
    editorElements.apiKeyInput.value = stored.apiKey;
  }
  if (stored.clientId) {
    editorElements.clientIdInput.value = stored.clientId;
  }
}

function saveCredentials() {
  const apiKey = editorElements.apiKeyInput.value.trim();
  const clientId = editorElements.clientIdInput.value.trim();
  if (!apiKey || !clientId) {
    setStatus('Enter both an API key and OAuth client ID before saving.', 'error');
    return;
  }
  localStorage.setItem('markdown-editor-drive-credentials', JSON.stringify({ apiKey, clientId }));
  setStatus('Saved Google API credentials locally.', 'success');
}

function showDriveError(error) {
  console.error(error);
  const message = typeof error === 'string' ? error : error?.result?.error?.message || error?.message || 'Unable to complete the Google Drive request.';
  editorElements.dialogAlert.textContent = message;
  editorElements.dialogAlert.hidden = false;
  setStatus(message, 'error');
}

function clearDriveError() {
  editorElements.dialogAlert.hidden = true;
  editorElements.dialogAlert.textContent = '';
}

function updateDriveButtons(isSignedIn) {
  const disabled = !isSignedIn;
  editorElements.driveOpenButton.disabled = disabled;
  editorElements.driveSaveButton.disabled = disabled;
  editorElements.driveSaveAsButton.disabled = disabled;
  editorElements.driveSignInButton.hidden = isSignedIn;
  editorElements.driveSignOutButton.hidden = !isSignedIn;
  editorElements.driveStatus.textContent = isSignedIn ? 'Connected to Google Drive' : 'Not connected';
}

function getCredentials() {
  const apiKey = editorElements.apiKeyInput.value.trim();
  const clientId = editorElements.clientIdInput.value.trim();
  if (!apiKey || !clientId) {
    throw new Error('Provide both an API key and OAuth client ID to connect to Google Drive.');
  }
  return { apiKey, clientId };
}

async function initializeGapiClient() {
  if (!gapiReady) {
    await waitForGapi();
  }
  if (gapiInitPromise) {
    return gapiInitPromise;
  }

  try {
    const { apiKey, clientId } = getCredentials();
    gapiInitPromise = new Promise((resolve, reject) => {
      gapi.load('client:auth2', async () => {
        try {
          await gapi.client.init({
            apiKey,
            clientId,
            discoveryDocs,
            scope: scopes
          });
          googleAuthInstance = gapi.auth2.getAuthInstance();
          googleAuthInstance.isSignedIn.listen(updateSigninStatus);
          updateSigninStatus(googleAuthInstance.isSignedIn.get());
          resolve();
        } catch (error) {
          gapiInitPromise = null;
          reject(error);
        }
      });
    });
    await gapiInitPromise;
    return;
  } catch (error) {
    gapiInitPromise = null;
    throw error;
  }
}

function updateSigninStatus(isSignedIn) {
  updateDriveButtons(isSignedIn);
  if (isSignedIn) {
    setStatus('Signed in to Google Drive.', 'success');
  } else {
    setStatus('Signed out of Google Drive.');
  }
}

async function waitForGapi() {
  if (gapiReady) {
    return;
  }
  await new Promise((resolve) => {
    const check = () => {
      if (gapiReady) {
        resolve();
      } else {
        window.setTimeout(check, 100);
      }
    };
    check();
  });
}

async function ensureDriveAccess() {
  await initializeGapiClient();
  if (!googleAuthInstance) {
    throw new Error('Unable to initialize Google authentication.');
  }
  if (!googleAuthInstance.isSignedIn.get()) {
    await googleAuthInstance.signIn();
  }
}

async function signInToGoogle() {
  try {
    await ensureDriveAccess();
  } catch (error) {
    showDriveError(error);
  }
}

function signOutOfGoogle() {
  if (googleAuthInstance) {
    googleAuthInstance.signOut();
  }
}

async function refreshDriveFileList() {
  clearDriveError();
  try {
    await ensureDriveAccess();
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
  try {
    await ensureDriveAccess();
    const response = await gapi.client.drive.files.get({ fileId, alt: 'media' });
    const content = response.body || response.result || '';
    editorElements.textarea.value = content;
    renderPreview(content);
    updateCounts(content);
    currentFileId = fileId;
    currentFileName = name;
    isDirty = false;
    updateFileIndicator();
    localStorage.setItem('markdown-editor-content', content);
    localStorage.setItem('markdown-editor-current-file', JSON.stringify({ id: fileId, name }));
    setStatus(`Loaded ${name} from Google Drive.`, 'success');
  } catch (error) {
    showDriveError(error);
  }
}

async function saveToDrive(forceNew = false) {
  clearDriveError();
  try {
    await ensureDriveAccess();
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
    const content = editorElements.textarea.value;
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
  restoreLastFile();
};

document.addEventListener('keydown', (event) => {
  if (event.target === editorElements.textarea && event.ctrlKey) {
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
