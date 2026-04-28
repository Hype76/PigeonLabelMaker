const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const { autoUpdater } = require("electron-updater");
const { spawn } = require("child_process");
const path = require("path");
const readline = require("readline");

const APP_NAME = "Pigeon Label Maker";
let mainWindow = null;

class BackendClient {
  constructor() {
    this.process = null;
    this.nextId = 1;
    this.pending = new Map();
    this.stopping = false;
  }

  start() {
    if (this.process) {
      return;
    }
    this.stopping = false;

    const cwd = app.isPackaged ? process.resourcesPath : path.resolve(__dirname, "..");
    const backendPath = app.isPackaged
      ? path.join(process.resourcesPath, "python", "backend.exe")
      : "python";
    const backendArgs = app.isPackaged
      ? []
      : ["-m", "pigeon_label_maker.backend_service"];

    this.process = spawn(backendPath, backendArgs, {
      cwd,
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true,
    });

    const output = readline.createInterface({ input: this.process.stdout });
    output.on("line", (line) => {
      let message;
      try {
        message = JSON.parse(line);
      } catch (error) {
        console.error("Backend JSON parse error", error, line);
        return;
      }
      const pending = this.pending.get(message.id);
      if (!pending) {
        return;
      }
      this.pending.delete(message.id);
      if (message.ok) {
        pending.resolve(message.result);
      } else {
        const error = new Error(message.error || "Backend request failed");
        error.traceback = message.traceback || "";
        pending.reject(error);
      }
    });

    this.process.stderr.on("data", (chunk) => {
      const text = chunk.toString().trim();
      if (text) {
        console.error("[backend]", text);
      }
    });

    this.process.on("exit", (code, signal) => {
      console.error(`Backend exited: code=${code}, signal=${signal}`);

      for (const pending of this.pending.values()) {
        pending.reject(new Error("Backend crashed"));
      }

      this.pending.clear();
      this.process = null;

      if (!this.stopping) {
        setTimeout(() => {
          console.log("Restarting backend...");
          this.start();
        }, 1000);
      }
    });
  }

  request(command, params = {}) {
    if (!this.process) {
      this.start();
    }
    return new Promise((resolve, reject) => {
      const id = this.nextId++;
      this.pending.set(id, { resolve, reject });
      this.process.stdin.write(JSON.stringify({ id, command, params }) + "\n");
    });
  }

  stop() {
    if (!this.process) {
      return;
    }
    this.stopping = true;
    this.process.kill();
    this.process = null;
  }
}

const backend = new BackendClient();

function appTitle() {
  return `${APP_NAME} v${app.getVersion()}`;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1500,
    height: 960,
    minWidth: 1180,
    minHeight: 760,
    icon: path.join(__dirname, "icon.png"),
    backgroundColor: "#d9d0c2",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.removeMenu();
  mainWindow.setTitle(appTitle());
  mainWindow.loadFile(path.join(__dirname, "index.html"));
}

function sendUpdateStatus(payload) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }
  mainWindow.webContents.send("update:status", payload);
}

function formatUpdateError(error) {
  const message = String(error?.message || error || "");

  if (message.includes("No published versions on GitHub")) {
    return "No update release published yet";
  }

  if (message.includes("Cannot find channel")) {
    return "Update metadata is missing from the release";
  }

  if (message.includes("net::ERR_INTERNET_DISCONNECTED")) {
    return "No internet connection";
  }

  return message || "Update check failed";
}

function setupAutoUpdater() {
  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on("checking-for-update", () => {
    sendUpdateStatus({ status: "checking", message: "Checking for updates..." });
  });

  autoUpdater.on("update-available", (info) => {
    sendUpdateStatus({
      status: "available",
      version: info.version,
      message: `Update ${info.version} available`,
    });
  });

  autoUpdater.on("update-not-available", () => {
    sendUpdateStatus({ status: "none", message: "App is up to date" });
  });

  autoUpdater.on("download-progress", (progress) => {
    sendUpdateStatus({
      status: "downloading",
      percent: Math.round(progress.percent || 0),
      message: `Downloading update ${Math.round(progress.percent || 0)}%`,
    });
  });

  autoUpdater.on("update-downloaded", (info) => {
    sendUpdateStatus({
      status: "downloaded",
      version: info.version,
      message: `Update ${info.version} ready to install`,
    });
  });

  autoUpdater.on("error", (error) => {
    sendUpdateStatus({
      status: "error",
      message: formatUpdateError(error),
    });
  });
}

ipcMain.handle("backend:request", async (_event, command, params) => {
  return backend.request(command, params);
});

ipcMain.handle("dialog:open-image", async () => {
  const result = await dialog.showOpenDialog({
    title: "Choose Image",
    filters: [{ name: "Images", extensions: ["png", "jpg", "jpeg", "bmp", "gif", "tif", "tiff"] }],
    properties: ["openFile"],
  });
  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }
  return result.filePaths[0];
});

ipcMain.handle("dialog:save-png", async () => {
  const result = await dialog.showSaveDialog({
    title: "Export PNG",
    defaultPath: "label.png",
    filters: [{ name: "PNG Image", extensions: ["png"] }],
  });
  return result.canceled ? null : result.filePath;
});

ipcMain.handle("app:version", async () => {
  return app.getVersion();
});

ipcMain.handle("update:check", async () => {
  if (!app.isPackaged) {
    const result = { status: "dev", message: "Updates work in the installed app only" };
    sendUpdateStatus(result);
    return result;
  }
  try {
    await autoUpdater.checkForUpdates();
    return { status: "checking", message: "Checking for updates..." };
  } catch (error) {
    const result = { status: "error", message: formatUpdateError(error) };
    sendUpdateStatus(result);
    return result;
  }
});

ipcMain.handle("update:download", async () => {
  if (!app.isPackaged) {
    const result = { status: "dev", message: "Updates work in the installed app only" };
    sendUpdateStatus(result);
    return result;
  }
  try {
    await autoUpdater.downloadUpdate();
    return { status: "downloading", message: "Downloading update..." };
  } catch (error) {
    const result = { status: "error", message: formatUpdateError(error) };
    sendUpdateStatus(result);
    return result;
  }
});

ipcMain.handle("update:install", async () => {
  if (!app.isPackaged) {
    const result = { status: "dev", message: "Updates work in the installed app only" };
    sendUpdateStatus(result);
    return result;
  }
  autoUpdater.quitAndInstall(false, true);
  return { status: "installing", message: "Installing update..." };
});

app.setName(APP_NAME);
setupAutoUpdater();

app.whenReady().then(() => {
  createWindow();
  if (app.isPackaged) {
    setTimeout(() => {
      autoUpdater.checkForUpdates().catch((error) => {
        sendUpdateStatus({
          status: "error",
          message: formatUpdateError(error),
        });
      });
    }, 4000);
  }
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  backend.stop();
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("before-quit", () => {
  backend.stop();
});
