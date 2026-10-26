/**
 * =====================================================================
 * SENTINEL STUDIO™ — ELECTRON STANDALONE DESKTOP WORKSTATION
 * Enterprise Capability Hub
 * Architect & Developer: Patrick Jay F. Tanap, CPA, CICA
 * Main Process & Multi-Session Partition Manager
 * =====================================================================
 */

const { app, BrowserWindow, ipcMain, dialog, session } = require('electron');
const path = require('path');
const fs = require('fs');

// Hardware Acceleration, WebGPU, and sRGB Profile Alignment
app.commandLine.appendSwitch('force-color-profile', 'srgb');
app.commandLine.appendSwitch('enable-unsafe-webgpu');
app.commandLine.appendSwitch('enable-features', 'Vulkan,WebGPU');
app.commandLine.appendSwitch('ignore-gpu-blocklist');

const PERSISTENT_SESSION_PARTITION = 'persist:sentinel_ai_workspace';

function configurePartitionHeaders(targetSession) {
  const filter = { urls: ['*://*/*'] };
  targetSession.webRequest.onHeadersReceived(filter, (details, callback) => {
    const responseHeaders = { ...details.responseHeaders };
    delete responseHeaders['x-frame-options'];
    delete responseHeaders['X-Frame-Options'];

    Object.keys(responseHeaders).forEach((key) => {
      if (key.toLowerCase() === 'content-security-policy') {
        const headerValues = responseHeaders[key] || [];
        responseHeaders[key] = headerValues.map((csp) => {
          return csp
            .replace(/frame-ancestors\s+[^;]+(;|$)/gi, '')
            .replace(/frame-src\s+[^;]+(;|$)/gi, "frame-src * https: data: blob: 'unsafe-inline'; ");
        });
      }
    });
    callback({ cancel: false, responseHeaders });
  });
}

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1200,
    minHeight: 768,
    backgroundColor: '#06234D',
    title: 'Sentinel Studio — Enterprise Capability Hub',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      webviewTag: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  mainWindow.setMenuBarVisibility(false);

  configurePartitionHeaders(session.defaultSession);
  const aiWorkspaceSession = session.fromPartition(PERSISTENT_SESSION_PARTITION);
  configurePartitionHeaders(aiWorkspaceSession);

  // Directly load the single-file HTML compiled by viteSingleFile
  mainWindow.loadFile(path.join(__dirname, 'dist', 'index.html'));

  return mainWindow;
}

function setupIpcHandlers() {
  ipcMain.handle('dialog:select-directory', async () => {
    const result = await dialog.showOpenDialog({
      properties: ['openDirectory', 'createDirectory'],
      title: 'Select Sentinel Workspace Mirror Directory',
    });
    if (result.canceled || !result.filePaths || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  });

  ipcMain.handle('fs:read-file', async (_event, filePath) => {
    return await fs.promises.readFile(filePath, 'utf-8');
  });

  ipcMain.handle('fs:write-file', async (_event, filePath, content) => {
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) await fs.promises.mkdir(dir, { recursive: true });
    await fs.promises.writeFile(filePath, content, 'utf-8');
    return { success: true, filePath };
  });

  ipcMain.handle('fs:save-binary', async (_event, filePath, data) => {
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) await fs.promises.mkdir(dir, { recursive: true });
    let buffer;
    if (Buffer.isBuffer(data)) buffer = data;
    else if (data instanceof ArrayBuffer) buffer = Buffer.from(data);
    else if (ArrayBuffer.isView(data)) buffer = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
    else if (typeof data === 'string' && data.startsWith('data:')) {
      buffer = Buffer.from(data.split(',')[1], 'base64');
    } else {
      buffer = Buffer.from(data, 'base64');
    }
    await fs.promises.writeFile(filePath, buffer);
    return { success: true, filePath };
  });

  ipcMain.handle('app:get-version', () => app.getVersion());
}

app.whenReady().then(() => {
  setupIpcHandlers();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});