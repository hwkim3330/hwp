const { app, BrowserWindow, Menu, Tray, nativeImage, shell, ipcMain } = require("electron");
const path = require("path");
const { spawn } = require("child_process");
const http = require("http");
const si = require("systeminformation");

const APP_URL = process.env.OFFICE_AGENT_APP_URL || "http://127.0.0.1:8765";
const PYTHON_CMD = process.env.OFFICE_AGENT_PYTHON || "python3";
const SERVER_ENTRY = path.resolve(__dirname, "..", "app.py");

let mainWindow = null;
let monitorWindow = null;
let visionWindow = null;
let tray = null;
let serverProcess = null;
let monitorTimer = null;
let hardwareInfo = {
  gpuLabel: "Unknown",
  gpuLoad: null,
  npuLabel: "Unavailable",
};
let previousSample = null;

function monitorMood(stats) {
  const stress = Math.max(stats.cpu, stats.mem);
  if (stress >= 90) {
    return { emoji: "🔥", face: "x_x", label: "critical" };
  }
  if (stress >= 75) {
    return { emoji: "😵", face: ">_<", label: "busy" };
  }
  if (stress >= 55) {
    return { emoji: "😼", face: "^_^", label: "active" };
  }
  if (stress >= 30) {
    return { emoji: "🙂", face: "o_o", label: "steady" };
  }
  return { emoji: "😴", face: "-_-", label: "idle" };
}

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
    width: 380,
    height: 560,
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

function createVisionWindow() {
  visionWindow = new BrowserWindow({
    width: 980,
    height: 760,
    show: false,
    title: "Vision Lab",
    backgroundColor: "#f5ecdf",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      sandbox: true,
    },
  });

  visionWindow.loadFile(path.join(__dirname, "vision.html"));
  visionWindow.on("closed", () => {
    visionWindow = null;
  });
}

async function fetchStats() {
  const [load, memory, battery, fsStats, networkStats] = await Promise.all([
    si.currentLoad(),
    si.mem(),
    si.battery(),
    si.fsStats(),
    si.networkStats(),
  ]);

  const cpu = Math.round(load.currentLoad);
  const mem = Math.round((memory.active / memory.total) * 100);
  const batt = battery.hasBattery ? Math.round(battery.percent) : null;
  const now = Date.now();
  const network = Array.isArray(networkStats) ? networkStats.find((item) => item.operstate === "up") || networkStats[0] : null;
  let netRxPerSec = 0;
  let netTxPerSec = 0;
  let diskReadPerSec = 0;
  let diskWritePerSec = 0;
  if (previousSample && now > previousSample.ts) {
    const seconds = (now - previousSample.ts) / 1000;
    if (seconds > 0) {
      if (network && previousSample.network) {
        netRxPerSec = Math.max(0, (network.rx_bytes - previousSample.network.rx_bytes) / seconds);
        netTxPerSec = Math.max(0, (network.tx_bytes - previousSample.network.tx_bytes) / seconds);
      }
      if (fsStats && previousSample.fs) {
        diskReadPerSec = Math.max(0, (fsStats.rx - previousSample.fs.rx) / seconds);
        diskWritePerSec = Math.max(0, (fsStats.wx - previousSample.fs.wx) / seconds);
      }
    }
  }
  previousSample = {
    ts: now,
    network: network ? { rx_bytes: network.rx_bytes, tx_bytes: network.tx_bytes } : null,
    fs: fsStats ? { rx: fsStats.rx, wx: fsStats.wx } : null,
  };
  return {
    cpu,
    mem,
    batt,
    loadAvg: Number(load.avgLoad || 0).toFixed(2),
    batteryCycleCount: battery.cycleCount ?? null,
    isCharging: Boolean(battery.isCharging),
    acConnected: Boolean(battery.acConnected),
    netRxPerSec: Math.round(netRxPerSec),
    netTxPerSec: Math.round(netTxPerSec),
    diskReadPerSec: Math.round(diskReadPerSec),
    diskWritePerSec: Math.round(diskWritePerSec),
    cpuCores: Array.isArray(load.cpus) ? load.cpus.map((item) => Math.round(item.load || 0)) : [],
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
    const mood = monitorMood(stats);
    const title = `${mood.emoji} ${stats.cpu}% ${stats.mem}%${stats.batt === null ? "" : ` ${stats.batt}%`}`;
    if (tray) {
      tray.setTitle(title);
      tray.setToolTip(`Office Agent Staff\n${mood.label.toUpperCase()}  CPU ${stats.cpu}%  MEM ${stats.mem}%${stats.batt === null ? "" : `  BAT ${stats.batt}%`}`);
    }
    if (monitorWindow && !monitorWindow.isDestroyed()) {
      monitorWindow.webContents.send("system-stats", { ...stats, mood });
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

function openVisionWindow() {
  if (!visionWindow || visionWindow.isDestroyed()) {
    createVisionWindow();
  }
  visionWindow.show();
  visionWindow.focus();
}

function createTray() {
  tray = new Tray(createTrayImage());
  tray.setTitle("SYS --");
  const contextMenu = Menu.buildFromTemplate([
    { label: "Open Office Agent", click: () => mainWindow?.show() },
    { label: "Toggle System Monitor", click: () => toggleMonitorWindow() },
    { label: "Open Vision Lab", click: () => openVisionWindow() },
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
  createVisionWindow();
  createTray();
  await updateMonitor();
  monitorTimer = setInterval(updateMonitor, 2000);
}

const VISION_RESOURCES = {
  headPointerGuide: "https://support.apple.com/en-mz/guide/mac-help/mchlb2d4782b/mac",
  displaysGuide: "https://support.apple.com/en-is/guide/mac-help/mchlb5f905a1/mac",
  externalCameraGuide: "https://support.apple.com/en-al/guide/mac-help/mchl034033f4/mac",
  mbp2021Specs: "https://support.apple.com/en-us/111901",
  mbp2024Specs: "https://support.apple.com/en-us/121554",
  lidAngleSensor: "https://github.com/samhenrigold/LidAngleSensor",
};

ipcMain.handle("vision:open-resource", async (_event, resource) => {
  const url = VISION_RESOURCES[resource];
  if (!url) {
    throw new Error(`unknown vision resource: ${resource}`);
  }
  await shell.openExternal(url);
  return { ok: true };
});

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
