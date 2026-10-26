/**
 * =====================================================================
 * SENTINEL STUDIO™ — ELECTRON STANDALONE DESKTOP WORKSTATION
 * Enterprise Capability Hub (ERI × EPEx)
 * Architect & Developer: Patrick Jay F. Tanap, CPA, CICA
 * Main Process & Multi-Session Partition Manager
 * =====================================================================
 */

const { app, BrowserWindow, ipcMain, dialog, session } = require('electron');
const path = require('path');
const fs = require('fs');

// 1. Hardware Acceleration & WebGPU CLI flags before app.ready
app.commandLine.appendSwitch('enable-unsafe-webgpu');
app.commandLine.appendSwitch('enable-features', 'Vulkan,WebGPU');
app.commandLine.appendSwitch('ignore-gpu-blocklist');

// Persistent partition key for multi-AI sessions
const PERSISTENT_SESSION_PARTITION = 'persist:sentinel_ai_workspace';

/**
 * Configure onHeadersReceived to strip X-Frame-Options and Content-Security-Policy
 * so that subscription web applications (Copilot, ChatGPT, Gemini) render cleanly
 * inside native <webview> partitions without CORS or framing refusals.
 */
function configurePartitionHeaders(targetSession) {
  const filter = { urls: ['*://*/*'] };

  targetSession.webRequest.onHeadersReceived(filter, (details, callback) => {
    const responseHeaders = { ...details.responseHeaders };

    // Strip X-Frame-Options in all casing variants
    delete responseHeaders['x-frame-options'];
    delete responseHeaders['X-Frame-Options'];

    // Sanitize Content-Security-Policy to eliminate frame-ancestors restrictions
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
  // 2. Native BrowserWindow Initialization
  const mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1200,
    minHeight: 768,
    backgroundColor: '#06234D',
    title: 'Sentinel Studio — Enterprise Capability Hub (ERI × EPEx)',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      webviewTag: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  // Remove default menu for clean enterprise look
  mainWindow.setMenuBarVisibility(false);

  // 3. Persistent Session Partition & Header Stripping
  configurePartitionHeaders(session.defaultSession);
  const aiWorkspaceSession = session.fromPartition(PERSISTENT_SESSION_PARTITION);
  configurePartitionHeaders(aiWorkspaceSession);

  // Load application entry point
  const devUrl = process.env.ELECTRON_START_URL;
  const distHtml = path.join(__dirname, 'dist', 'index.html');
  const rootHtml = path.join(__dirname, 'index.html');

  if (devUrl) {
    mainWindow.loadURL(devUrl).catch(() => {
      if (fs.existsSync(distHtml)) {
        mainWindow.loadFile(distHtml);
      } else {
        mainWindow.loadFile(rootHtml);
      }
    });
  } else if (fs.existsSync(distHtml)) {
    mainWindow.loadFile(distHtml);
  } else {
    mainWindow.loadFile(rootHtml);
  }

  return mainWindow;
}

// 4. IPC (ipcMain) Native Handlers
function setupIpcHandlers() {
  // Directory Selection
  ipcMain.handle('dialog:select-directory', async () => {
    const result = await dialog.showOpenDialog({
      properties: ['openDirectory', 'createDirectory'],
      title: 'Select Sentinel Workspace Mirror Directory',
    });
    if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
      return null;
    }
    return result.filePaths[0];
  });

  // UTF-8 File Read
  ipcMain.handle('fs:read-file', async (_event, filePath) => {
    try {
      return await fs.promises.readFile(filePath, 'utf-8');
    } catch (err) {
      console.error('[IPC fs:read-file Error]', err);
      throw err;
    }
  });

  // UTF-8 Text File Write
  ipcMain.handle('fs:write-file', async (_event, filePath, content) => {
    try {
      const dir = path.dirname(filePath);
      if (!fs.existsSync(dir)) {
        await fs.promises.mkdir(dir, { recursive: true });
      }
      await fs.promises.writeFile(filePath, content, 'utf-8');
      return { success: true, filePath };
    } catch (err) {
      console.error('[IPC fs:write-file Error]', err);
      throw err;
    }
  });

  // Binary File Write (Direct native export for DOCX, XLSX, PPTX)
  ipcMain.handle('fs:save-binary', async (_event, filePath, data) => {
    try {
      const dir = path.dirname(filePath);
      if (!fs.existsSync(dir)) {
        await fs.promises.mkdir(dir, { recursive: true });
      }

      let buffer;
      if (Buffer.isBuffer(data)) {
        buffer = data;
      } else if (data instanceof ArrayBuffer) {
        buffer = Buffer.from(data);
      } else if (ArrayBuffer.isView(data)) {
        buffer = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
      } else if (typeof data === 'string') {
        if (data.startsWith('data:')) {
          const base64Data = data.split(',')[1];
          buffer = Buffer.from(base64Data, 'base64');
        } else {
          buffer = Buffer.from(data, 'base64');
        }
      } else {
        buffer = Buffer.from(data);
      }

      await fs.promises.writeFile(filePath, buffer);
      return { success: true, filePath };
    } catch (err) {
      console.error('[IPC fs:save-binary Error]', err);
      throw err;
    }
  });

  // Application Version
  ipcMain.handle('app:get-version', () => {
    return app.getVersion();
  });
}

// App lifecycle
app.whenReady().then(() => {
  setupIpcHandlers();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
