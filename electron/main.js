const { app, BrowserWindow, Menu, Tray, nativeImage, shell } = require("electron");
const path = require("path");
const { spawn } = require("child_process");
const http = require("http");
const si = require("systeminformation");

const APP_URL = process.env.OFFICE_AGENT_APP_URL || "http://127.0.0.1:8765";
const PYTHON_CMD = process.env.OFFICE_AGENT_PYTHON || "python3";
const SERVER_ENTRY = path.resolve(__dirname, "..", "app.py");

let mainWindow = null;
let monitorWindow = null;
let tray = null;
let serverProcess = null;
let monitorTimer = null;
let hardwareInfo = {
  gpuLabel: "Unknown",
  gpuLoad: null,
  npuLabel: "Unavailable",
};

function createTrayImage() {
  const image = nativeImage.createEmpty();
  image.setTemplateImage(true);
  return image;
}

function waitForServer(url, timeoutMs = 30000) {
  const start = Date.now();
  return new Promise((resolve, reject) => {
    const tryConnect = () => {
      const req = http.get(`${url}/healthz`, (res) => {
        res.resume();
        if (res.statusCode === 200) {
          resolve();
        } else {
          retry();
        }
      });
      req.on("error", retry);
      req.setTimeout(1500, () => {
        req.destroy();
        retry();
      });
    };

    const retry = () => {
      if (Date.now() - start > timeoutMs) {
        reject(new Error("Office Agent server did not start in time"));
        return;
      }
      setTimeout(tryConnect, 500);
    };

    tryConnect();
  });
}

function startServer() {
  if (serverProcess) {
    return Promise.resolve();
  }

  serverProcess = spawn(PYTHON_CMD, [SERVER_ENTRY], {
    cwd: path.resolve(__dirname, ".."),
    env: process.env,
    stdio: "ignore",
  });

  serverProcess.on("exit", () => {
    serverProcess = null;
  });

  return waitForServer(APP_URL);
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1480,
    height: 980,
    minWidth: 1180,
    minHeight: 760,
    title: "Office Agent Staff",
    backgroundColor: "#f4efe6",
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      sandbox: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  mainWindow.loadURL(APP_URL);
  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

function createMonitorWindow() {
  monitorWindow = new BrowserWindow({
    width: 320,
    height: 220,
    show: false,
    frame: false,
    resizable: false,
    movable: true,
    minimizable: false,
    maximizable: false,
    fullscreenable: false,
    title: "System Monitor",
    backgroundColor: "#f6f0e7",
    alwaysOnTop: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      sandbox: true,
    },
  });

  monitorWindow.loadFile(path.join(__dirname, "monitor.html"));
  monitorWindow.on("blur", () => {
    if (monitorWindow && monitorWindow.isVisible()) {
      monitorWindow.hide();
    }
  });
}

async function fetchStats() {
  const [load, memory, battery] = await Promise.all([
    si.currentLoad(),
    si.mem(),
    si.battery(),
  ]);

  const cpu = Math.round(load.currentLoad);
  const mem = Math.round((memory.active / memory.total) * 100);
  const batt = battery.hasBattery ? Math.round(battery.percent) : null;
  return {
    cpu,
    mem,
    batt,
    gpuLabel: hardwareInfo.gpuLabel,
    gpuLoad: hardwareInfo.gpuLoad,
    npuLabel: hardwareInfo.npuLabel,
  };
}

async function loadHardwareInfo() {
  try {
    const graphics = await si.graphics();
    const primaryGpu = graphics.controllers.find((controller) => controller.model) || graphics.controllers[0];
    if (primaryGpu) {
      hardwareInfo.gpuLabel = primaryGpu.cores
        ? `${primaryGpu.model} (${primaryGpu.cores} cores)`
        : primaryGpu.model;
    }
  } catch (_error) {
    hardwareInfo.gpuLabel = "Unavailable";
  }
}

async function updateMonitor() {
  try {
    const stats = await fetchStats();
    const title = `CPU ${stats.cpu}% MEM ${stats.mem}%${stats.batt === null ? "" : ` BAT ${stats.batt}%`}`;
    if (tray) {
      tray.setTitle(title);
      tray.setToolTip(`Office Agent Staff\n${title}`);
    }
    if (monitorWindow && !monitorWindow.isDestroyed()) {
      monitorWindow.webContents.send("system-stats", stats);
    }
  } catch (error) {
    if (tray) {
      tray.setTitle("SYS --");
    }
  }
}

function toggleMonitorWindow() {
  if (!monitorWindow) {
    return;
  }
  if (monitorWindow.isVisible()) {
    monitorWindow.hide();
    return;
  }
  const trayBounds = tray.getBounds();
  const windowBounds = monitorWindow.getBounds();
  monitorWindow.setPosition(Math.round(trayBounds.x - windowBounds.width / 2), Math.round(trayBounds.y + 28), false);
  monitorWindow.show();
  monitorWindow.focus();
}

function createTray() {
  tray = new Tray(createTrayImage());
  tray.setTitle("SYS --");
  const contextMenu = Menu.buildFromTemplate([
    { label: "Open Office Agent", click: () => mainWindow?.show() },
    { label: "Toggle System Monitor", click: () => toggleMonitorWindow() },
    { type: "separator" },
    { label: "Open in Browser", click: () => shell.openExternal(APP_URL) },
    { type: "separator" },
    { label: "Quit", role: "quit" },
  ]);
  tray.setContextMenu(contextMenu);
  tray.on("click", () => toggleMonitorWindow());
}

async function bootstrap() {
  await startServer();
  await loadHardwareInfo();
  createMainWindow();
  createMonitorWindow();
  createTray();
  await updateMonitor();
  monitorTimer = setInterval(updateMonitor, 2000);
}

app.whenReady().then(async () => {
  await bootstrap();
  app.on("activate", () => {
    if (!mainWindow) {
      createMainWindow();
    }
  });
});

app.on("window-all-closed", (event) => {
  event.preventDefault();
});

app.on("before-quit", () => {
  if (monitorTimer) {
    clearInterval(monitorTimer);
  }
  if (serverProcess) {
    serverProcess.kill();
  }
});
