// SheetCraft AI desktop: starts the bundled backend (which also serves the screens) on this
// computer and shows it in a window. The Anthropic API key is asked for once and kept here.
const { app, BrowserWindow, Menu, dialog, ipcMain, net, safeStorage, shell } = require("electron");
const { spawn } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const netSocket = require("node:net");

// A fixed address keeps the browser storage (saved workbooks) the same from one launch to the next.
const PORTS = [47615, 47616, 47617, 47618, 47619];
const settingsFile = () => path.join(app.getPath("userData"), "settings.json");
const dataDir = () => path.join(app.getPath("userData"), "data");
const logFile = () => path.join(app.getPath("userData"), "server.log");

let server = null;
let mainWindow = null;
let quitting = false;
let started = false; // the main window has opened; until then, closing the key window must not quit

function readSettings() {
  try {
    return JSON.parse(fs.readFileSync(settingsFile(), "utf8"));
  } catch {
    return {};
  }
}

function writeSettings(s) {
  fs.mkdirSync(app.getPath("userData"), { recursive: true });
  fs.writeFileSync(settingsFile(), JSON.stringify(s, null, 2));
}

function getKey() {
  const s = readSettings();
  if (!s.key) return null;
  try {
    return s.encrypted ? safeStorage.decryptString(Buffer.from(s.key, "base64")) : s.key;
  } catch {
    return null;
  }
}

function saveKey(key) {
  const s = readSettings();
  // Encrypted with the system keychain (Keychain on Mac, DPAPI on Windows) when it's available.
  if (safeStorage.isEncryptionAvailable()) {
    s.key = safeStorage.encryptString(key).toString("base64");
    s.encrypted = true;
  } else {
    s.key = key;
    s.encrypted = false;
  }
  writeSettings(s);
}

function portFree(port) {
  return new Promise((resolve) => {
    const srv = netSocket.createServer();
    srv.once("error", () => resolve(false));
    srv.once("listening", () => srv.close(() => resolve(true)));
    srv.listen(port, "127.0.0.1");
  });
}

async function pickPort() {
  const last = readSettings().port;
  for (const p of last ? [last, ...PORTS.filter((x) => x !== last)] : PORTS) if (await portFree(p)) return p;
  throw new Error(`Ports ${PORTS[0]}-${PORTS.at(-1)} are all in use. Close the program using them and try again.`);
}

/** Where LibreOffice usually lives, so PDF export works when it's installed. */
function withOfficeOnPath(env) {
  const extra =
    process.platform === "darwin"
      ? ["/Applications/LibreOffice.app/Contents/MacOS"]
      : process.platform === "win32"
        ? ["C:\\Program Files\\LibreOffice\\program", "C:\\Program Files (x86)\\LibreOffice\\program"]
        : [];
  const key = Object.keys(env).find((k) => k.toLowerCase() === "path") || "PATH";
  return { ...env, [key]: [env[key], ...extra].filter(Boolean).join(path.delimiter) };
}

function serverCommand(port) {
  const web = app.isPackaged ? path.join(process.resourcesPath, "web") : path.join(__dirname, "..", "frontend", "out");
  const args = ["--port", String(port), "--web", web, "--data", dataDir(), "--exit-with-parent"];
  if (app.isPackaged) {
    const exe = process.platform === "win32" ? "sheetcraft-server.exe" : "sheetcraft-server";
    return { cmd: path.join(process.resourcesPath, "server", exe), args, cwd: path.join(process.resourcesPath, "server") };
  }
  // Running from the source folder: use Python and the backend next to this folder.
  const backend = path.join(__dirname, "..", "backend");
  return { cmd: process.platform === "win32" ? "python" : "python3", args: [path.join(backend, "desktop_server.py"), ...args], cwd: backend };
}

async function startServer() {
  const port = await pickPort();
  writeSettings({ ...readSettings(), port });
  const { cmd, args, cwd } = serverCommand(port);
  fs.mkdirSync(app.getPath("userData"), { recursive: true });
  const log = fs.openSync(logFile(), "w");
  server = spawn(cmd, args, {
    cwd,
    env: withOfficeOnPath({ ...process.env, ANTHROPIC_API_KEY: getKey() }),
    // An open stdin tells the engine we're alive; it stops by itself if this app ever dies.
    stdio: ["pipe", log, log],
    windowsHide: true,
  });
  let exited = null;
  server.on("exit", (code) => {
    exited = code;
    server = null;
    if (!quitting) {
      dialog.showErrorBox("SheetCraft AI stopped", `The SheetCraft engine stopped unexpectedly (code ${code}).\n\nDetails are in ${logFile()}`);
      app.quit();
    }
  });
  const url = `http://127.0.0.1:${port}`;
  for (let i = 0; i < 120; i++) {
    if (exited !== null) throw new Error(`The SheetCraft engine didn't start. Details are in ${logFile()}`);
    try {
      const res = await net.fetch(`${url}/health`);
      if (res.ok) return url;
    } catch {}
    await new Promise((r) => setTimeout(r, 250));
  }
  throw new Error(`The SheetCraft engine took too long to start. Details are in ${logFile()}`);
}

function stopServer() {
  if (server) server.kill();
  server = null;
}

/** First launch (or "Change API key"): a small window that asks for the key. Resolves once one is saved. */
function askForKey(parent) {
  return new Promise((resolve) => {
    const win = new BrowserWindow({
      width: 520,
      height: 470,
      resizable: false,
      parent: parent ?? undefined,
      modal: !!parent,
      title: "SheetCraft AI",
      autoHideMenuBar: true,
      webPreferences: { preload: path.join(__dirname, "preload.js") },
    });
    let saved = false;
    const onSave = async (_e, key) => {
      key = String(key || "").trim();
      if (!/^sk-ant-[\w-]{20,}$/.test(key)) return { ok: false, message: "That doesn't look like an Anthropic API key. It starts with sk-ant-." };
      // Check it with Anthropic; if we can't reach them, save it anyway and let the app report problems.
      try {
        const res = await net.fetch("https://api.anthropic.com/v1/models?limit=1", {
          headers: { "x-api-key": key, "anthropic-version": "2023-06-01" },
        });
        if (res.status === 401 || res.status === 403) return { ok: false, message: "Anthropic didn't accept that key. Check it and try again." };
      } catch {}
      saveKey(key);
      saved = true;
      win.close();
      return { ok: true };
    };
    ipcMain.handle("save-key", onSave);
    win.on("closed", () => {
      ipcMain.removeHandler("save-key");
      resolve(saved);
    });
    win.webContents.setWindowOpenHandler(({ url }) => {
      shell.openExternal(url);
      return { action: "deny" };
    });
    win.loadFile(path.join(__dirname, "setup.html"));
  });
}

function buildMenu() {
  const changeKey = {
    label: "Change API key…",
    click: async () => {
      if (await askForKey(mainWindow)) {
        // The engine reads the key when it starts, so restart it.
        quitting = true;
        stopServer();
        quitting = false;
        try {
          const url = await startServer();
          mainWindow?.loadURL(url);
        } catch (e) {
          dialog.showErrorBox("SheetCraft AI", e.message);
        }
      }
    },
  };
  const openData = { label: "Open data folder", click: () => shell.openPath(app.getPath("userData")) };
  const isMac = process.platform === "darwin";
  const template = [
    ...(isMac ? [{ label: app.name, submenu: [{ role: "about" }, { type: "separator" }, changeKey, openData, { type: "separator" }, { role: "hide" }, { role: "hideOthers" }, { role: "unhide" }, { type: "separator" }, { role: "quit" }] }] : []),
    { label: "File", submenu: [...(isMac ? [] : [changeKey, openData, { type: "separator" }]), isMac ? { role: "close" } : { role: "quit" }] },
    { role: "editMenu" },
    { label: "View", submenu: [{ role: "reload" }, { role: "toggleDevTools" }, { type: "separator" }, { role: "resetZoom" }, { role: "zoomIn" }, { role: "zoomOut" }, { type: "separator" }, { role: "togglefullscreen" }] },
    { role: "windowMenu" },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

async function createWindow(url) {
  mainWindow = new BrowserWindow({ width: 1440, height: 900, minWidth: 800, minHeight: 560, title: "SheetCraft AI", show: false });
  const local = (u) => u.startsWith(url);
  // Links that leave the app (Google Sheets, help pages, checkout) open in the normal browser.
  mainWindow.webContents.setWindowOpenHandler(({ url: target }) => {
    if (!local(target)) shell.openExternal(target);
    return { action: "deny" };
  });
  mainWindow.webContents.on("will-navigate", (e, target) => {
    if (!local(target)) {
      e.preventDefault();
      shell.openExternal(target);
    }
  });
  mainWindow.once("ready-to-show", () => mainWindow.show());
  mainWindow.on("closed", () => (mainWindow = null));
  await mainWindow.loadURL(url);
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on("second-instance", () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });

  app.whenReady().then(async () => {
    buildMenu();
    if (!getKey() && !(await askForKey(null))) return app.quit();
    try {
      await createWindow(await startServer());
      started = true;
    } catch (e) {
      dialog.showErrorBox("SheetCraft AI", e.message);
      app.quit();
    }
  });

  app.on("activate", async () => {
    // Mac: clicking the dock icon after closing the window opens it again.
    if (!mainWindow && server) await createWindow(`http://127.0.0.1:${readSettings().port}`);
  });

  app.on("window-all-closed", () => {
    if (started && process.platform !== "darwin") app.quit();
  });

  app.on("before-quit", () => {
    quitting = true;
    stopServer();
  });
}
