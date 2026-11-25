import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { app, BrowserWindow, ipcMain, session, dialog } from 'electron';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join } from 'node:path';
import { createSharedModelStore } from './shared/shared-model-store.mjs';
import { modelAction } from './models.mjs';
import { createOfficeSession } from './office-session.mjs';

const root = dirname(fileURLToPath(import.meta.url));
const page = join(root, 'ui', 'offline.html');
const windows = new Map();
const contexts = new Map();
const requestedHost = (args) =>
  args.map((v) => /^--office=(word|excel|powerpoint)$/.exec(v)?.[1]).find(Boolean);
const store = createSharedModelStore();
app.commandLine.appendSwitch('disable-background-networking');
async function prefersEnglish() {
  return (await store.state().catch(() => undefined))?.uiLanguage === 'en';
}
async function openHost(host) {
  if (!host) {
    const english = await prefersEnglish();
    const result = dialog.showMessageBoxSync({
      type: 'question',
      message: english ? 'Choose the Office app you are using' : '选择正在使用的 Office 应用',
      buttons: ['Word', 'Excel', 'PowerPoint', english ? 'Cancel' : '取消'],
      cancelId: 3,
    });
    host = ['word', 'excel', 'powerpoint'][result];
    if (!host) return;
  }
  const existing = windows.get(host);
  if (existing) {
    if (existing.isMinimized()) existing.restore();
    existing.show();
    existing.focus();
    return;
  }
  const window = new BrowserWindow({
    width: 1000,
    height: 790,
    minWidth: 390,
    minHeight: 530,
    show: false,
    title: '三桥 AI · 本地工作台',
    webPreferences: {
      preload: join(root, 'preload.cjs'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      webviewTag: false,
      navigateOnDragDrop: false,
      additionalArguments: ['--trio-host=' + host],
    },
  });
  windows.set(host, window);
  const requests = new Map();
  const office = createOfficeSession(host, root, {
    openDeck:
      process.platform === 'darwin'
        ? (file) => promisify(execFile)('/usr/bin/open', ['-b', 'com.microsoft.Powerpoint', file])
        : undefined,
    saveDeck: async () => {
      const english = await prefersEnglish();
      const { filePath, canceled } = await dialog.showSaveDialog(window, {
        title: english ? 'Save and open in PowerPoint' : '保存并在 PowerPoint 中打开',
        defaultPath: english ? 'TrioBridge presentation.pptx' : '三桥演示稿.pptx',
        filters: [
          { name: english ? 'PowerPoint presentation' : 'PowerPoint 演示稿', extensions: ['pptx'] },
        ],
      });
      return canceled ? undefined : filePath;
    },
  });
  const id = window.webContents.id;
  contexts.set(id, { requests, office });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event) => event.preventDefault());
  window.webContents.on('will-attach-webview', (event) => event.preventDefault());
  window.once('ready-to-show', () => window.show());
  window.on('closed', () => {
    windows.delete(host);
    contexts.delete(id);
    for (const x of requests.values()) x.abort();
    office.close();
  });
  void window.loadFile(page);
}
function context(event) {
  const frame = event.senderFrame;
  if (
    !frame ||
    frame !== event.sender.mainFrame ||
    frame.url.split('#')[0] !== pathToFileURL(page).href
  )
    throw Error('无效的窗口。');
  const ctx = contexts.get(event.sender.id);
  if (!ctx) throw Error('窗口已关闭。');
  return ctx;
}
if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', (_e, args) => {
    void app.whenReady().then(() => openHost(requestedHost(args)));
  });
  void app.whenReady().then(() => {
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] },
      (_details, done) => done({ cancel: true }),
    );
    const allowCopy = (contents, permission) =>
      permission === 'clipboard-sanitized-write' &&
      contents &&
      contexts.has(contents.id) &&
      contents.getURL().split('#')[0] === pathToFileURL(page).href;
    session.defaultSession.setPermissionRequestHandler((contents, permission, done) =>
      done(Boolean(allowCopy(contents, permission))),
    );
    session.defaultSession.setPermissionCheckHandler((contents, permission) =>
      Boolean(allowCopy(contents, permission)),
    );
    ipcMain.handle('trio:model', async (event, id, action, payload) => {
      const { requests } = context(event);
      if (
        typeof id !== 'string' ||
        !/^[a-f0-9-]{36}$/.test(id) ||
        requests.size >= 12 ||
        requests.has(id)
      )
        throw Error('请求繁忙，请重试。');
      const controller = new AbortController();
      requests.set(id, controller);
      try {
        return await modelAction(store, action, payload, controller.signal);
      } finally {
        requests.delete(id);
      }
    });
    ipcMain.on('trio:cancel', (event, id) => context(event).requests.get(id)?.abort());
    ipcMain.handle('trio:office', (event, action, data) =>
      context(event).office.invoke(action, data),
    );
    openHost(requestedHost(process.argv));
    app.on('window-all-closed', () => app.quit());
  });
}
