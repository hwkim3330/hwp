const { app, BrowserWindow, Menu, Tray, nativeImage, shell, ipcMain, screen, globalShortcut } = require("electron");
const path = require("path");
const { spawn, execFile } = require("child_process");
const http = require("http");
const si = require("systeminformation");

const APP_URL = process.env.OFFICE_AGENT_APP_URL || "http://127.0.0.1:8765";
const PYTHON_CMD = process.env.OFFICE_AGENT_PYTHON || "python3";
const SERVER_ENTRY = path.resolve(__dirname, "..", "app.py");

let mainWindow = null;
let monitorWindow = null;
let visionWindow = null;
let cursorOverlayWindow = null;
let tray = null;
let serverProcess = null;
let monitorTimer = null;
let cursorTimer = null;
let cursorOverlayEnabled = false;
let captureStatus = { ok: true, message: "클립보드 캡처 대기" };
let hardwareInfo = {
  gpuLabel: "Unknown",
  gpuLoad: null,
  npuLabel: "Unavailable",
};
let previousSample = null;
const OPERATOR_PRESETS = {
  claude_fast: {
    label: "Claude Fast",
    text: "설명은 줄이고 바로 실행해. 필요한 수정만 적용하고 마지막에 검증 결과만 짧게 남겨.",
    submit: true,
  },
  codex_fast: {
    label: "Codex Fast",
    text: "분석은 짧게 하고 바로 구현해. 테스트나 검증까지 끝내고 핵심만 보고해.",
    submit: true,
  },
  hwp_manager: {
    label: "hwp Manager",
    text: "문서 초안을 구조화 블록 우선으로 만들고, 필요하면 검색 결과를 참고 링크와 함께 정리해.",
    submit: true,
  },
  continue_work: {
    label: "Continue",
    text: "중단한 작업 이어서 진행해. 불필요한 설명 없이 바로 처리해.",
    submit: true,
  },
};

function getVirtualDisplayBounds() {
  const displays = screen.getAllDisplays();
  const left = Math.min(...displays.map((display) => display.bounds.x));
  const top = Math.min(...displays.map((display) => display.bounds.y));
  const right = Math.max(...displays.map((display) => display.bounds.x + display.bounds.width));
  const bottom = Math.max(...displays.map((display) => display.bounds.y + display.bounds.height));
  return {
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  };
}

function monitorMood(stats) {
  const stress = Math.max(stats.cpu, stats.mem);
  if (stress >= 90) {
    return { avatar: "=^x_x^=", face: "x_x", label: "critical" };
  }
  if (stress >= 75) {
    return { avatar: "=^>_<^=", face: ">_<", label: "busy" };
  }
  if (stress >= 55) {
    return { avatar: "=^^_^=", face: "^_^", label: "active" };
  }
  if (stress >= 30) {
    return { avatar: "=^o_o^=", face: "o_o", label: "steady" };
  }
  return { avatar: "=^-_-^=", face: "-_-", label: "idle" };
}

function createTrayImage() {
  const image = nativeImage.createEmpty();
  image.setTemplateImage(true);
  return image;
}

function runAppleScript(script) {
  return new Promise((resolve, reject) => {
    execFile("osascript", ["-e", script], (error, stdout, stderr) => {
      if (error) {
        reject(new Error(stderr || error.message));
        return;
      }
      resolve(stdout);
    });
  });
}

function refocusPreviousApp() {
  const script = [
    'tell application "System Events"',
    "  key down command",
    "  key code 48",
    "  key up command",
    "end tell",
  ].join("\n");
  return runAppleScript(script);
}

function escapeAppleScriptText(value) {
  return String(value).replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function typeToFrontmostApp(text, submit = true) {
  const lines = [
    'tell application "System Events"',
    `  keystroke "${escapeAppleScriptText(text)}"`,
  ];
  if (submit) {
    lines.push("  key code 36");
  }
  lines.push("end tell");
  return runAppleScript(lines.join("\n"));
}

async function runOperatorPreset(presetId) {
  const preset = OPERATOR_PRESETS[presetId];
  if (!preset) {
    throw new Error(`unknown preset: ${presetId}`);
  }
  if (process.platform !== "darwin") {
    throw new Error("operator presets currently support macOS only");
  }
  if (monitorWindow && !monitorWindow.isDestroyed()) {
    monitorWindow.hide();
  }
  await new Promise((resolve) => setTimeout(resolve, 40));
  await refocusPreviousApp();
  await new Promise((resolve) => setTimeout(resolve, 160));
  await typeToFrontmostApp(preset.text, preset.submit !== false);
  return { preset: presetId, label: preset.label };
}

function captureScreenshotToClipboard() {
  return new Promise((resolve, reject) => {
    if (process.platform !== "darwin") {
      reject(new Error("clipboard screenshot currently supports macOS only"));
      return;
    }
    execFile("screencapture", ["-i", "-c"], (error, stdout, stderr) => {
      if (error) {
        reject(new Error(stderr || error.message));
        return;
      }
      resolve({ ok: true });
    });
  });
}

function requestJson(url, method = "GET", body = null, timeoutMs = 1500) {
  return new Promise((resolve, reject) => {
    const target = new URL(url);
    const req = http.request(
      {
        protocol: target.protocol,
        hostname: target.hostname,
        port: target.port,
        path: `${target.pathname}${target.search}`,
        method,
        headers: body ? { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(body) } : {},
      },
      (res) => {
        let raw = "";
        res.setEncoding("utf8");
        res.on("data", (chunk) => {
          raw += chunk;
        });
        res.on("end", () => {
          try {
            resolve(raw ? JSON.parse(raw) : {});
          } catch (error) {
            reject(error);
          }
        });
      },
    );
    req.on("error", reject);
    req.setTimeout(timeoutMs, () => {
      req.destroy(new Error("request timed out"));
    });
    if (body) {
      req.write(body);
    }
    req.end();
  });
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
        reject(new Error("hwp server did not start in time"));
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
    title: "hwp",
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

function updateCursorOverlayFrame() {
  const cursor = screen.getCursorScreenPoint();
  if (cursorOverlayWindow && !cursorOverlayWindow.isDestroyed()) {
    const bounds = cursorOverlayWindow.getBounds();
    cursorOverlayWindow.webContents.send("cursor-overlay-tick", {
      cursor: { x: cursor.x - bounds.x, y: cursor.y - bounds.y },
      enabled: cursorOverlayEnabled,
    });
  }
  if (monitorWindow && !monitorWindow.isDestroyed()) {
    monitorWindow.webContents.send("cursor-overlay-state", { enabled: cursorOverlayEnabled, cursor });
  }
}

function createCursorOverlayWindow() {
  const display = getVirtualDisplayBounds();
  cursorOverlayWindow = new BrowserWindow({
    x: display.x,
    y: display.y,
    width: display.width,
    height: display.height,
    show: false,
    frame: false,
    transparent: true,
    focusable: false,
    skipTaskbar: true,
    hasShadow: false,
    resizable: false,
    movable: false,
    fullscreenable: false,
    title: "Cursor Overlay",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      sandbox: true,
    },
  });
  cursorOverlayWindow.setBounds(display);
  cursorOverlayWindow.setAlwaysOnTop(true, "screen-saver");
  cursorOverlayWindow.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
  cursorOverlayWindow.setIgnoreMouseEvents(true, { forward: true });
  cursorOverlayWindow.loadFile(path.join(__dirname, "cursor-overlay.html"));
  cursorOverlayWindow.on("closed", () => {
    cursorOverlayWindow = null;
  });
}

async function fetchStats() {
  const [load, memory, battery, fsStats, networkStats, ollamaPs] = await Promise.all([
    si.currentLoad(),
    si.mem(),
    si.battery(),
    si.fsStats(),
    si.networkStats(),
    requestJson("http://127.0.0.1:11434/api/ps").catch(() => ({ models: [] })),
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
    cursor: screen.getCursorScreenPoint(),
    cursorOverlayEnabled,
    captureStatus,
    llm: Array.isArray(ollamaPs.models) && ollamaPs.models.length > 0
      ? {
          active: true,
          name: ollamaPs.models[0].name || "unknown",
          processor: ollamaPs.models[0].processor || "unknown",
          size: ollamaPs.models[0].size || "",
        }
      : {
          active: false,
          name: "gemma4:latest",
          processor: "idle",
          size: "",
        },
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
    const title = `${mood.face} ${stats.cpu}% ${stats.mem}%${stats.batt === null ? "" : ` ${stats.batt}%`}`;
    if (tray) {
      tray.setTitle(title);
      tray.setToolTip(`hwp\n${mood.label.toUpperCase()}  CPU ${stats.cpu}%  MEM ${stats.mem}%${stats.batt === null ? "" : `  BAT ${stats.batt}%`}`);
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

function toggleCursorOverlay(forceState) {
  cursorOverlayEnabled = typeof forceState === "boolean" ? forceState : !cursorOverlayEnabled;
  if (!cursorOverlayWindow || cursorOverlayWindow.isDestroyed()) {
    createCursorOverlayWindow();
  }
  cursorOverlayWindow.setBounds(getVirtualDisplayBounds());
  if (cursorOverlayEnabled) {
    cursorOverlayWindow.show();
  } else if (cursorOverlayWindow && !cursorOverlayWindow.isDestroyed()) {
    cursorOverlayWindow.hide();
  }
  if (monitorWindow && !monitorWindow.isDestroyed()) {
    monitorWindow.webContents.send("cursor-overlay-state", { enabled: cursorOverlayEnabled });
  }
  return cursorOverlayEnabled;
}

function createTray() {
  tray = new Tray(createTrayImage());
  tray.setTitle("SYS --");
  const contextMenu = Menu.buildFromTemplate([
    { label: "Open hwp", click: () => mainWindow?.show() },
    { label: "Toggle System Monitor", click: () => toggleMonitorWindow() },
    { label: "Clipboard Screenshot", click: () => ipcCaptureScreenshot() },
    { label: "Toggle Cursor Spotlight", click: () => toggleCursorOverlay() },
    { label: "Open Vision Lab", click: () => openVisionWindow() },
    { type: "separator" },
    { label: "Open in Browser", click: () => shell.openExternal(APP_URL) },
    { type: "separator" },
    { label: "Quit", role: "quit" },
  ]);
  tray.setContextMenu(contextMenu);
  tray.on("click", () => toggleMonitorWindow());
}

async function ipcCaptureScreenshot() {
  try {
    captureStatus = { ok: true, message: "캡처 대기 중" };
    await captureScreenshotToClipboard();
    captureStatus = { ok: true, message: "클립보드에 캡처 완료" };
  } catch (error) {
    captureStatus = { ok: false, message: String(error.message || error) };
    throw error;
  }
  return captureStatus;
}

async function bootstrap() {
  await startServer();
  await loadHardwareInfo();
  createMainWindow();
  createMonitorWindow();
  createVisionWindow();
  createCursorOverlayWindow();
  createTray();
  globalShortcut.register("CommandOrControl+Shift+1", () => {
    ipcCaptureScreenshot().catch(() => {});
  });
  await updateMonitor();
  monitorTimer = setInterval(updateMonitor, 2000);
  cursorTimer = setInterval(updateCursorOverlayFrame, 33);
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

ipcMain.handle("cursor-overlay:toggle", async (_event, forceState) => {
  return { ok: true, enabled: toggleCursorOverlay(forceState) };
});

ipcMain.handle("cursor-overlay:state", async () => {
  return { ok: true, enabled: cursorOverlayEnabled };
});

ipcMain.handle("operator:run-preset", async (_event, presetId) => {
  const result = await runOperatorPreset(presetId);
  return { ok: true, result };
});

ipcMain.handle("capture:screenshot", async () => {
  const result = await ipcCaptureScreenshot();
  return { ok: true, result };
});

app.whenReady().then(async () => {
  await bootstrap();
  const syncCursorOverlayBounds = () => {
    if (cursorOverlayWindow && !cursorOverlayWindow.isDestroyed()) {
      cursorOverlayWindow.setBounds(getVirtualDisplayBounds());
    }
  };
  screen.on("display-added", syncCursorOverlayBounds);
  screen.on("display-removed", syncCursorOverlayBounds);
  screen.on("display-metrics-changed", syncCursorOverlayBounds);
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
  if (cursorTimer) {
    clearInterval(cursorTimer);
  }
  if (serverProcess) {
    serverProcess.kill();
  }
});
