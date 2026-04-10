import init, { HwpDocument } from "https://cdn.jsdelivr.net/npm/@rhwp/core@0.6.1/rhwp.js";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";

const WASM_URL = "https://cdn.jsdelivr.net/npm/@rhwp/core@0.6.1/rhwp_bg.wasm";

const elements = {
  topTabWriter: document.querySelector("#top-tab-writer"),
  topTabNotes: document.querySelector("#top-tab-notes"),
  topTabSheet: document.querySelector("#top-tab-sheet"),
  topTabSlides: document.querySelector("#top-tab-slides"),
  topbarRoute: document.querySelector("#topbar-route"),
  topbarStatus: document.querySelector("#topbar-status"),
  settingsToggle: document.querySelector("#settings-toggle"),
  focusToggle: document.querySelector("#focus-toggle"),
  tabWriter: document.querySelector("#tab-writer"),
  tabNotes: document.querySelector("#tab-notes"),
  tabSheet: document.querySelector("#tab-sheet"),
  tabSlides: document.querySelector("#tab-slides"),
  insertHeading: document.querySelector("#insert-heading"),
  insertParagraph: document.querySelector("#insert-paragraph"),
  insertTable: document.querySelector("#insert-table"),
  addWriterParagraph: document.querySelector("#add-writer-paragraph"),
  workspaceWriter: document.querySelector("#workspace-writer"),
  workspaceNotes: document.querySelector("#workspace-notes"),
  workspaceSheet: document.querySelector("#workspace-sheet"),
  workspaceSlides: document.querySelector("#workspace-slides"),
  newDocument: document.querySelector("#new-document"),
  exportWorkspace: document.querySelector("#export-workspace"),
  exportDocument: document.querySelector("#export-document"),
  exportDocx: document.querySelector("#export-docx"),
  fileInput: document.querySelector("#file-input"),
  newNote: document.querySelector("#new-note"),
  exportNote: document.querySelector("#export-note"),
  notesPad: document.querySelector("#notes-pad"),
  addSheetRow: document.querySelector("#add-sheet-row"),
  addSheetColumn: document.querySelector("#add-sheet-column"),
  addSheetTotals: document.querySelector("#add-sheet-totals"),
  resetSheet: document.querySelector("#reset-sheet"),
  exportSheet: document.querySelector("#export-sheet"),
  exportXlsx: document.querySelector("#export-xlsx"),
  sheetGrid: document.querySelector("#sheet-grid"),
  addSlide: document.querySelector("#add-slide"),
  generateSlides: document.querySelector("#generate-slides"),
  exportSlides: document.querySelector("#export-slides"),
  exportPptx: document.querySelector("#export-pptx"),
  slidesDeck: document.querySelector("#slides-deck"),
  docMeta: document.querySelector("#doc-meta"),
  outlineBox: document.querySelector("#outline-box"),
  recentCommands: document.querySelector("#recent-commands"),
  memoryMeta: document.querySelector("#memory-meta"),
  memoryRecall: document.querySelector("#memory-recall"),
  projectPill: document.querySelector("#project-pill"),
  projectFocusState: document.querySelector("#project-focus-state"),
  projectName: document.querySelector("#project-name"),
  projectGoal: document.querySelector("#project-goal"),
  statusBox: document.querySelector("#status-box"),
  renderBadge: document.querySelector("#render-badge"),
  liveActivityTitle: document.querySelector("#live-activity-title"),
  liveActivityRoute: document.querySelector("#live-activity-route"),
  liveActivityDetail: document.querySelector("#live-activity-detail"),
  liveActivityShortcut: document.querySelector("#live-activity-shortcut"),
  liveActivityProgress: document.querySelector("#live-activity-progress"),
  commandRouteHint: document.querySelector("#command-route-hint"),
  agentRuntime: document.querySelector("#agent-runtime"),
  dashboardNow: document.querySelector("#dashboard-now"),
  dashboardCapture: document.querySelector("#dashboard-capture"),
  dashboardPlan: document.querySelector("#dashboard-plan"),
  modeHint: document.querySelector("#mode-hint"),
  pages: document.querySelector("#pages"),
  writerEditor: document.querySelector("#writer-editor"),
  promptInput: document.querySelector("#prompt-input"),
  saveMemory: document.querySelector("#save-memory"),
  runAgent: document.querySelector("#run-agent"),
  reply: document.querySelector("#agent-reply"),
  planBox: document.querySelector("#plan-box"),
  plannerMeta: document.querySelector("#planner-meta"),
  capLlm: document.querySelector("#cap-llm"),
  capLlmMeta: document.querySelector("#cap-llm-meta"),
  capHwpforge: document.querySelector("#cap-hwpforge"),
  capHwpforgeMeta: document.querySelector("#cap-hwpforge-meta"),
  capComputerUse: document.querySelector("#cap-computer-use"),
  capComputerUseMeta: document.querySelector("#cap-computer-use-meta"),
  toolRegistry: document.querySelector("#tool-registry"),
  permissionRegistry: document.querySelector("#permission-registry"),
  sessionMeta: document.querySelector("#session-meta"),
  sessionLog: document.querySelector("#session-log"),
  actionResearchNote: document.querySelector("#action-research-note"),
  actionResearchComplete: document.querySelector("#action-research-complete"),
  actionMinutes: document.querySelector("#action-minutes"),
  actionReport: document.querySelector("#action-report"),
  actionSlides: document.querySelector("#action-slides"),
  searchEnabled: document.querySelector("#search-enabled"),
  modelProfile: document.querySelector("#model-profile"),
  workflowHint: document.querySelector("#workflow-hint"),
  editorEngine: document.querySelector("#editor-engine"),
  openOnlyOffice: document.querySelector("#open-onlyoffice"),
  engineMeta: document.querySelector("#engine-meta"),
  onlyofficeSessions: document.querySelector("#onlyoffice-sessions"),
  computerUseRunNext: document.querySelector("#computer-use-run-next"),
  computerUseRunAll: document.querySelector("#computer-use-run-all"),
  computerUseProgressLabel: document.querySelector("#computer-use-progress-label"),
  computerUseProgressMeta: document.querySelector("#computer-use-progress-meta"),
  computerUseProgressBar: document.querySelector("#computer-use-progress-bar"),
  computerUseMeta: document.querySelector("#computer-use-meta"),
  computerUsePlan: document.querySelector("#computer-use-plan"),
  computerUseSessions: document.querySelector("#computer-use-sessions"),
  searchResults: document.querySelector("#search-results"),
  openBrowser: document.querySelector("#open-browser"),
  openFinder: document.querySelector("#open-finder"),
  openMonitor: document.querySelector("#open-monitor"),
  systemActionLog: document.querySelector("#system-action-log"),
  appModal: document.querySelector("#app-modal"),
  closeAppModal: document.querySelector("#close-app-modal"),
  settingsModelProfile: document.querySelector("#settings-model-profile"),
  settingsSearchDefault: document.querySelector("#settings-search-default"),
  settingsSkipOnboarding: document.querySelector("#settings-skip-onboarding"),
  templateCards: [...document.querySelectorAll(".template-card")],
  promptChips: [...document.querySelectorAll(".prompt-chip")],
};

const SHEET_COLUMNS = ["항목", "담당", "상태", "기한", "우선순위", "비고"];

const state = {
  doc: null,
  fileName: "untitled.hwp",
  ready: false,
  mode: "writer",
  noteText: "",
  sheetColumns: [...SHEET_COLUMNS],
  sheetRows: [],
  slides: [],
  currentComputerUsePlan: null,
  currentComputerUseSessionId: "",
  computerUseBusy: false,
  liveRoute: "auto",
  commandHistory: [],
  autosavedAt: 0,
  project: {
    name: "Untitled Project",
    goal: "",
    focusMode: false,
  },
  preferences: {
    onboardingDismissed: false,
    defaultSearchEnabled: false,
    defaultModelProfile: "balanced",
  },
};

const STORAGE_KEY = "hwp-state-v1";
const MAX_COMMAND_HISTORY = 12;
let writerEditorSyncTimer = 0;
let memoryRecallTimer = 0;

function installMeasureTextWidth() {
  let ctx = null;
  let lastFont = "";
  globalThis.measureTextWidth = (font, text) => {
    if (!ctx) {
      ctx = document.createElement("canvas").getContext("2d");
    }
    if (font !== lastFont) {
      ctx.font = font;
      lastFont = font;
    }
    return ctx.measureText(text).width;
  };
}

function setStatus(message, extra = "") {
  elements.statusBox.textContent = extra ? `${message}\n${extra}` : message;
  if (elements.topbarStatus) {
    elements.topbarStatus.textContent = extra || message;
  }
  if (elements.liveActivityTitle) {
    elements.liveActivityTitle.textContent = message;
  }
  if (elements.liveActivityDetail) {
    elements.liveActivityDetail.textContent = extra || message;
  }
}

function updateProjectUi() {
  const name = String(state.project?.name || "").trim() || "Untitled Project";
  const goal = String(state.project?.goal || "").trim();
  if (elements.projectPill) {
    elements.projectPill.textContent = name;
  }
  if (elements.projectName && elements.projectName.value !== name) {
    elements.projectName.value = name;
  }
  if (elements.projectGoal && elements.projectGoal.value !== goal) {
    elements.projectGoal.value = goal;
  }
  if (elements.projectFocusState) {
    elements.projectFocusState.textContent = state.project?.focusMode ? "Focus" : "Standard";
  }
  document.body.classList.toggle("focus-mode", Boolean(state.project?.focusMode));
  if (elements.dashboardPlan && goal && !elements.dashboardPlan.textContent.trim()) {
    elements.dashboardPlan.textContent = goal;
  }
}

function openAppModal() {
  elements.appModal?.classList.remove("hidden");
}

function closeAppModal() {
  elements.appModal?.classList.add("hidden");
}

function syncPreferencesUi() {
  if (elements.settingsModelProfile) {
    elements.settingsModelProfile.value = state.preferences.defaultModelProfile || "balanced";
  }
  if (elements.settingsSearchDefault) {
    elements.settingsSearchDefault.checked = Boolean(state.preferences.defaultSearchEnabled);
  }
  if (elements.settingsSkipOnboarding) {
    elements.settingsSkipOnboarding.checked = Boolean(state.preferences.onboardingDismissed);
  }
  if (elements.modelProfile) {
    elements.modelProfile.value = state.preferences.defaultModelProfile || "balanced";
  }
  if (elements.searchEnabled) {
    elements.searchEnabled.checked = Boolean(state.preferences.defaultSearchEnabled);
  }
}

function setFocusMode(force) {
  const next = typeof force === "boolean" ? force : !state.project.focusMode;
  state.project.focusMode = next;
  updateProjectUi();
  persistWorkspace();
  setStatus(next ? "집중 모드를 켰습니다." : "표준 작업 모드로 돌아왔습니다.");
}

function formatTimestamp(ts) {
  if (!ts) {
    return "-";
  }
  return new Date(ts).toLocaleTimeString("ko-KR", { hour12: false });
}

function setBadge(message) {
  elements.renderBadge.textContent = message;
  if (elements.liveActivityProgress) {
    elements.liveActivityProgress.textContent = message;
  }
}

function setRuntimeBadge(message) {
  if (elements.agentRuntime) {
    elements.agentRuntime.textContent = message;
  }
  if (elements.dashboardNow) {
    elements.dashboardNow.textContent = `런타임 상태: ${message}`;
  }
}

function setMode(mode) {
  state.mode = mode;
  const map = {
    writer: [elements.tabWriter, elements.workspaceWriter],
    notes: [elements.tabNotes, elements.workspaceNotes],
    sheet: [elements.tabSheet, elements.workspaceSheet],
    slides: [elements.tabSlides, elements.workspaceSlides],
  };

  [elements.tabWriter, elements.tabNotes, elements.tabSheet, elements.tabSlides].forEach((tab) =>
    tab.classList.remove("active"),
  );
  [elements.topTabWriter, elements.topTabNotes, elements.topTabSheet, elements.topTabSlides].forEach((tab) =>
    tab?.classList.remove("active"),
  );
  [elements.workspaceWriter, elements.workspaceNotes, elements.workspaceSheet, elements.workspaceSlides].forEach((panel) =>
    panel.classList.remove("active"),
  );
  map[mode][0].classList.add("active");
  map[mode][1].classList.add("active");
  const topMap = {
    writer: elements.topTabWriter,
    notes: elements.topTabNotes,
    sheet: elements.topTabSheet,
    slides: elements.topTabSlides,
  };
  topMap[mode]?.classList.add("active");
  elements.modeHint.textContent = `현재 대상: ${mode[0].toUpperCase()}${mode.slice(1)}`;
  persistWorkspace();
}

function setWorkflowHint(message) {
  if (elements.workflowHint) {
    elements.workflowHint.textContent = message;
  }
  if (elements.dashboardCapture) {
    elements.dashboardCapture.textContent = message;
  }
  if (!String(state.project?.goal || "").trim() && elements.projectGoal) {
    elements.projectGoal.value = message.slice(0, 140);
    state.project.goal = elements.projectGoal.value;
    updateProjectUi();
  }
}

function applyTemplate(template) {
  if (template === "research") {
    state.project.name = "연구노트 패키지";
    state.project.goal = "조사, 참고 링크, 문서 초안까지 한 번에 정리";
    if (elements.promptInput) {
      elements.promptInput.value = "연구노트를 완성형으로 작성해줘. 배경, 핵심 질문, 조사 요약, 참고 링크, 다음 액션을 포함해.";
    }
    if (elements.searchEnabled) {
      elements.searchEnabled.checked = true;
    }
    if (elements.modelProfile) {
      elements.modelProfile.value = "deep";
    }
  }
  if (template === "report") {
    state.project.name = "업무 보고 패키지";
    state.project.goal = "요약, 핵심 내용, 일정표 중심 보고서 작성";
    if (elements.promptInput) {
      elements.promptInput.value = "보고서 형식으로 요약, 핵심 내용, 일정 표를 포함해 작성해줘.";
    }
    if (elements.searchEnabled) {
      elements.searchEnabled.checked = false;
    }
    if (elements.modelProfile) {
      elements.modelProfile.value = "balanced";
    }
  }
  if (template === "slides") {
    state.project.name = "발표 패키지";
    state.project.goal = "문서와 슬라이드 초안을 같이 준비";
    if (elements.promptInput) {
      elements.promptInput.value = "현재 내용을 발표용 슬라이드 초안과 발표 메모까지 이어지게 정리해줘.";
    }
    if (elements.searchEnabled) {
      elements.searchEnabled.checked = true;
    }
    if (elements.modelProfile) {
      elements.modelProfile.value = "balanced";
    }
  }
  updateProjectUi();
  previewAgentRoute();
  persistWorkspace();
  closeAppModal();
}

function setLiveRoute(route, detail = "") {
  state.liveRoute = route;
  const routeLabel = route === "computer_use" ? "Browser" : route === "document" ? "Document" : "Auto";
  if (elements.liveActivityRoute) {
    elements.liveActivityRoute.textContent = routeLabel;
  }
  if (elements.commandRouteHint) {
    elements.commandRouteHint.textContent = routeLabel;
  }
  if (elements.topbarRoute) {
    elements.topbarRoute.textContent = routeLabel;
  }
  if (elements.liveActivityShortcut) {
    elements.liveActivityShortcut.textContent =
      route === "computer_use"
        ? "명령 제출 시 브라우저 계획과 자동 진행을 실행합니다."
        : "Cmd/Ctrl + Enter로 바로 실행";
  }
  if (detail && elements.liveActivityDetail) {
    elements.liveActivityDetail.textContent = detail;
  }
}

function previewAgentRoute() {
  const prompt = String(elements.promptInput?.value || "").trim();
  const route = detectAgentRoute(prompt);
  if (!prompt) {
    if (elements.commandRouteHint) {
      elements.commandRouteHint.textContent = "Document";
    }
    window.clearTimeout(memoryRecallTimer);
    memoryRecallTimer = window.setTimeout(() => {
      refreshMemoryRecall("");
    }, 120);
    return;
  }
  if (elements.commandRouteHint) {
    elements.commandRouteHint.textContent = route === "computer_use" ? "Browser" : "Document";
  }
  window.clearTimeout(memoryRecallTimer);
  memoryRecallTimer = window.setTimeout(() => {
    refreshMemoryRecall(prompt);
  }, 180);
}

async function applyStartupCommandFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const prompt = String(params.get("prompt") || "").trim();
  const route = String(params.get("route") || "").trim();
  const autorun = params.get("autorun") === "1";
  if (!prompt && !route) {
    return;
  }
  if (elements.promptInput && prompt) {
    elements.promptInput.value = prompt;
  }
  previewAgentRoute();
  if (route === "computer_use") {
    setLiveRoute("computer_use", prompt);
    if (autorun && prompt) {
      await planComputerUse(prompt, { autorun: true });
    }
  } else if (autorun && prompt) {
    await runAgent();
  }
  window.history.replaceState({}, "", "/");
}

function setEngineMeta(message) {
  if (elements.engineMeta) {
    elements.engineMeta.textContent = message;
  }
  if (elements.dashboardPlan) {
    elements.dashboardPlan.textContent = message;
  }
}

function renderOnlyOfficeSessions(items) {
  if (!elements.onlyofficeSessions) {
    return;
  }
  if (!Array.isArray(items) || items.length === 0) {
    elements.onlyofficeSessions.textContent = "아직 생성된 세션이 없습니다.";
    return;
  }
  elements.onlyofficeSessions.innerHTML = items
    .map(
      (item) => `
        <div class="session-event">
          <strong>${escapeHtml(item.title || item.id)}</strong>
          <span>${escapeHtml((item.extension || "").toUpperCase())} · ${escapeHtml(item.mode || "-")} · ${escapeHtml(formatSessionTime(item.created_at))}</span>
          <code>${escapeHtml(item.share_url || "")}</code>
          <div class="session-link-row">
            <button class="secondary onlyoffice-open" data-url="${encodeURIComponent(item.launch_url || "")}">열기</button>
            <button class="secondary onlyoffice-copy" data-url="${encodeURIComponent(item.share_url || "")}">링크 복사</button>
          </div>
        </div>
      `,
    )
    .join("");
  elements.onlyofficeSessions.querySelectorAll(".onlyoffice-open").forEach((button) => {
    button.addEventListener("click", () => {
      const url = decodeURIComponent(button.dataset.url || "");
      if (url) {
        window.open(url, "_blank", "noopener,noreferrer");
      }
    });
  });
  elements.onlyofficeSessions.querySelectorAll(".onlyoffice-copy").forEach((button) => {
    button.addEventListener("click", async () => {
      const url = decodeURIComponent(button.dataset.url || "");
      if (!url) {
        return;
      }
      try {
        await navigator.clipboard.writeText(url);
        setEngineMeta(`공유 링크 복사 완료: ${url}`);
      } catch (error) {
        setEngineMeta(String(error.message || error));
      }
    });
  });
}

function setComputerUsePresetGoal(goal) {
  if (elements.promptInput) {
    elements.promptInput.value = goal;
  }
  if (elements.computerUseMeta) {
    elements.computerUseMeta.textContent = `브라우저 목표 준비: ${goal}`;
  }
  previewAgentRoute();
  setLiveRoute("computer_use", goal);
}

function renderCommandHistory() {
  if (!elements.recentCommands) {
    return;
  }
  if (!Array.isArray(state.commandHistory) || state.commandHistory.length === 0) {
    elements.recentCommands.textContent = "아직 최근 작업이 없습니다.";
    return;
  }
  elements.recentCommands.innerHTML = state.commandHistory
    .slice(0, 6)
    .map(
      (item, index) => `
        <div class="session-event">
          <strong>${escapeHtml(item.route === "computer_use" ? "브라우저 작업" : "문서 작업")}</strong>
          <span>${escapeHtml(formatTimestamp(item.ts))}</span>
          <code>${escapeHtml(item.prompt || "")}</code>
          <div class="session-link-row">
            <button class="secondary recent-command-run" data-index="${index}">다시 실행</button>
          </div>
        </div>
      `,
    )
    .join("");
  elements.recentCommands.querySelectorAll(".recent-command-run").forEach((button) => {
    button.addEventListener("click", async () => {
      const item = state.commandHistory[Number(button.dataset.index || "-1")];
      if (!item) {
        return;
      }
      if (elements.promptInput) {
        elements.promptInput.value = item.prompt || "";
      }
      previewAgentRoute();
      await runAgent();
    });
  });
}

function renderMemoryRecall(items) {
  if (!elements.memoryRecall) {
    return;
  }
  if (!Array.isArray(items) || items.length === 0) {
    elements.memoryRecall.textContent = "아직 회수된 기억이 없습니다.";
    return;
  }
  elements.memoryRecall.innerHTML = items
    .map(
      (item) => `
        <div class="session-event">
          <strong>${escapeHtml(item.title || item.kind || "memory")}</strong>
          <span>${escapeHtml(item.kind || "-")} · ${escapeHtml(formatTimestamp((item.ts || 0) * 1000))}${item.pinned ? " · pinned" : ""}</span>
          <code>${escapeHtml(item.text || "")}</code>
          <div class="session-link-row">
            <button class="secondary memory-use" data-title="${escapeHtml(item.title || "")}" data-text="${escapeHtml(item.text || "")}">프롬프트로</button>
            <button class="secondary memory-pin" data-id="${escapeHtml(item.id || "")}" data-pinned="${item.pinned ? "0" : "1"}">${item.pinned ? "고정 해제" : "고정"}</button>
            <button class="secondary memory-delete" data-id="${escapeHtml(item.id || "")}">삭제</button>
          </div>
        </div>
      `,
    )
    .join("");
  elements.memoryRecall.querySelectorAll(".memory-use").forEach((button) => {
    button.addEventListener("click", () => {
      const title = decodeHtml(button.dataset.title || "");
      const text = decodeHtml(button.dataset.text || "");
      if (elements.promptInput) {
        elements.promptInput.value = title ? `${title}\n${text}` : text;
      }
      previewAgentRoute();
      setStatus("기억을 프롬프트로 불러왔습니다.");
    });
  });
  elements.memoryRecall.querySelectorAll(".memory-pin").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = decodeHtml(button.dataset.id || "");
      const pinned = button.dataset.pinned === "1";
      try {
        const response = await fetch("/api/memory/pin", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id, pinned }),
        });
        const data = await response.json();
        if (!data.ok) {
          throw new Error(data.detail || data.error || "memory pin failed");
        }
        await refreshMemoryRecall(String(elements.promptInput?.value || "").trim());
      } catch (error) {
        elements.memoryRecall.textContent = String(error.message || error);
      }
    });
  });
  elements.memoryRecall.querySelectorAll(".memory-delete").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = decodeHtml(button.dataset.id || "");
      try {
        const response = await fetch("/api/memory/delete", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id }),
        });
        const data = await response.json();
        if (!data.ok) {
          throw new Error(data.detail || data.error || "memory delete failed");
        }
        await refreshMemoryRecall(String(elements.promptInput?.value || "").trim());
      } catch (error) {
        elements.memoryRecall.textContent = String(error.message || error);
      }
    });
  });
}

async function refreshMemoryRecall(query = "") {
  if (!elements.memoryRecall) {
    return;
  }
  try {
    const response = await fetch(`/api/memory/search?q=${encodeURIComponent(query)}`, { cache: "no-store" });
    const data = await response.json();
    if (!data.ok) {
      throw new Error("memory unavailable");
    }
    renderMemoryRecall(data.items || []);
    if (elements.memoryMeta) {
      elements.memoryMeta.textContent = query
        ? `현재 입력 기준 기억 ${data.items?.length || 0}개`
        : `최근 기억 ${data.items?.length || 0}개`;
    }
  } catch (error) {
    elements.memoryRecall.textContent = String(error.message || error);
  }
}

function rememberCommand(prompt, route) {
  const text = String(prompt || "").trim();
  if (!text) {
    return;
  }
  state.commandHistory = [
    { prompt: text, route, ts: Date.now() },
    ...state.commandHistory.filter((item) => item.prompt !== text),
  ].slice(0, MAX_COMMAND_HISTORY);
  renderCommandHistory();
}

function getWriterSnapshot() {
  const summary = getDocumentSummary();
  return {
    paragraphs: summary.paragraphs.map((item) => item.text).filter(Boolean),
  };
}

function formatComputerUseAction(action) {
  if (!action || typeof action !== "object") {
    return "";
  }
  if (action.type === "open_url") {
    return action.url || "";
  }
  if (action.type === "search_query") {
    return action.query || "";
  }
  if (action.type === "open_app") {
    return action.app || "";
  }
  return action.text || "";
}

function decodeHtml(text) {
  const textarea = document.createElement("textarea");
  textarea.innerHTML = text;
  return textarea.value;
}

function summarizeComputerUseProgress(session) {
  const total = Array.isArray(session?.actions) ? session.actions.length : Array.isArray(session?.plan?.actions) ? session.plan.actions.length : 0;
  const done = Number(session?.executed_steps ?? session?.history?.length ?? 0);
  const currentIndex = done < total ? done : Math.max(total - 1, 0);
  const percent = total > 0 ? Math.round((done / total) * 100) : 0;
  return { total, done, currentIndex, percent };
}

function setComputerUseProgress(session) {
  const progress = summarizeComputerUseProgress(session);
  if (elements.computerUseProgressLabel) {
    elements.computerUseProgressLabel.textContent =
      progress.total > 0 ? `${progress.done}/${progress.total} 단계 완료` : "대기 중";
  }
  if (elements.computerUseProgressMeta) {
    const status = session?.status ? `상태 ${session.status}` : "브라우저 작업 전";
    elements.computerUseProgressMeta.textContent =
      progress.total > 0 ? `${status} · 다음 단계 ${Math.min(progress.done + 1, progress.total)}` : "계획을 만들면 진행률과 다음 단계가 여기에 표시됩니다.";
  }
  if (elements.computerUseProgressBar) {
    elements.computerUseProgressBar.style.width = `${progress.percent}%`;
  }
  if (elements.computerUseRunNext) {
    elements.computerUseRunNext.disabled = state.computerUseBusy || progress.total === 0 || progress.done >= progress.total;
  }
  if (elements.computerUseRunAll) {
    elements.computerUseRunAll.disabled = state.computerUseBusy || progress.total === 0 || progress.done >= progress.total;
  }
}

function renderComputerUsePlan(session) {
  if (!elements.computerUsePlan) {
    return;
  }
  const plan = session?.plan || session;
  const sessionId = session?.id || state.currentComputerUseSessionId;
  if (!plan || !Array.isArray(plan.actions) || plan.actions.length === 0) {
    elements.computerUsePlan.textContent = "아직 생성된 브라우저 계획이 없습니다.";
    setComputerUseProgress(null);
    return;
  }
  const progress = summarizeComputerUseProgress({
    actions: plan.actions,
    executed_steps: session?.executed_steps ?? 0,
    status: session?.status,
  });
  elements.computerUsePlan.innerHTML = `
    <div class="session-event">
      <strong>${escapeHtml(plan.summary || "브라우저 작업 계획")}</strong>
      <span>${escapeHtml(plan.reply || "")}</span>
      <code>${escapeHtml(`세션 ${sessionId} · ${plan.meta?.planner || session?.status || "-"}`)}</code>
    </div>
    ${plan.actions
      .map(
        (action, index) => `
          <div class="computer-step${index < progress.done ? " is-done" : ""}${index === progress.currentIndex && progress.done < progress.total ? " is-current" : ""}">
            <div class="computer-step-top">
              <span class="computer-step-index">${index + 1}</span>
              <span class="computer-step-label">${escapeHtml(action.label || action.type || "step")}</span>
            </div>
            <div class="computer-step-meta">${escapeHtml(formatComputerUseAction(action))}</div>
          </div>
        `,
      )
      .join("")}
  `;
  setComputerUseProgress({
    actions: plan.actions,
    executed_steps: session?.executed_steps ?? 0,
    status: session?.status,
  });
}

function renderComputerUseSessions(items) {
  if (!elements.computerUseSessions) {
    return;
  }
  if (!Array.isArray(items) || items.length === 0) {
    elements.computerUseSessions.textContent = "아직 브라우저 세션이 없습니다.";
    return;
  }
  const latest = items[0];
  if (latest) {
    state.currentComputerUseSessionId = latest.id || "";
    state.currentComputerUsePlan = latest;
    renderComputerUsePlan(latest);
  }
  elements.computerUseSessions.innerHTML = items.slice(0, 4)
    .map(
      (item) => `
        <div class="session-event">
          <strong>${escapeHtml(item.goal || item.id)}</strong>
          <span>${escapeHtml(item.status || "planned")} · 실행 ${escapeHtml(String(item.executed_steps || 0))}회 · ${escapeHtml(formatSessionTime(item.created_at))}</span>
          <code>${escapeHtml(item.summary || "")}</code>
          <div class="session-link-row">
            <button class="secondary computer-use-resume" data-session-id="${escapeHtml(item.id)}">불러오기</button>
          </div>
        </div>
      `,
    )
    .join("");
  elements.computerUseSessions.querySelectorAll(".computer-use-resume").forEach((button) => {
    button.addEventListener("click", () => {
      const session = items.find((item) => item.id === (button.dataset.sessionId || ""));
      if (session) {
        state.currentComputerUseSessionId = session.id;
        state.currentComputerUsePlan = session;
        renderComputerUsePlan(session);
        if (elements.computerUseMeta) {
          elements.computerUseMeta.textContent = `세션 불러옴 · ${session.goal}`;
        }
      }
    });
  });
}

async function refreshAgentHealth() {
  try {
    const response = await fetch("/healthz", { cache: "no-store" });
    const health = await response.json();
    if (!health.ok) {
      throw new Error("health unavailable");
    }
    elements.capLlm.textContent = health.model || "unknown";
    elements.capLlmMeta.textContent = health.baseUrl || "LLM endpoint unavailable";
    if (health.hwpforge?.available) {
      elements.capHwpforge.textContent = "ready";
      elements.capHwpforgeMeta.textContent = health.hwpforge.command || "hwpforge";
    } else {
      elements.capHwpforge.textContent = "offline";
      elements.capHwpforgeMeta.textContent = health.hwpforge?.detail || "hwpforge unavailable";
    }
    if (elements.capComputerUse) {
      elements.capComputerUse.textContent = `${health.computerUse?.sessions ?? 0} sessions`;
    }
    if (elements.capComputerUseMeta) {
      elements.capComputerUseMeta.textContent = health.computerUse?.reference?.detail || "browser-use reference unavailable";
    }
    setEngineMeta(`ONLYOFFICE Docs: ${health.onlyoffice?.docsUrl || "http://127.0.0.1:8080"} | active sessions ${health.onlyoffice?.sessions ?? 0}`);
    setRuntimeBadge(health.hwpforge?.available ? "Local Ops Ready" : "Partial Local Mode");
    if (elements.dashboardNow) {
      elements.dashboardNow.textContent = `모델 ${health.model || "unknown"} · HWPX ${health.hwpforge?.available ? "ready" : "offline"} · OOXML ${health.onlyoffice?.sessions ?? 0} · Browser ${health.computerUse?.sessions ?? 0} · Memory ${health.memory?.items ?? 0}`;
    }
  } catch (error) {
    elements.capLlm.textContent = "offline";
    elements.capLlmMeta.textContent = "health check failed";
    elements.capHwpforge.textContent = "unknown";
    elements.capHwpforgeMeta.textContent = String(error.message || error);
    if (elements.capComputerUse) {
      elements.capComputerUse.textContent = "unknown";
    }
    if (elements.capComputerUseMeta) {
      elements.capComputerUseMeta.textContent = String(error.message || error);
    }
    setEngineMeta(String(error.message || error));
    setRuntimeBadge("Offline");
  }
}

function renderRegistryRows(target, items, emptyText) {
  if (!target) {
    return;
  }
  if (!Array.isArray(items) || items.length === 0) {
    target.textContent = emptyText;
    return;
  }
  target.innerHTML = items
    .map(
      (item) => `
        <div class="registry-row">
          <strong>${escapeHtml(item.label || item.id || "unknown")}</strong>
          <span>${escapeHtml(item.detail || "")}</span>
          <span class="registry-state">${escapeHtml(item.status || "unknown")}</span>
        </div>
      `,
    )
    .join("");
}

function renderSearchResults(results) {
  if (!elements.searchResults) {
    return;
  }
  if (!Array.isArray(results) || results.length === 0) {
    elements.searchResults.textContent = "검색 결과가 없습니다.";
    return;
  }
  elements.searchResults.innerHTML = results
    .map(
      (item) => `
        <div class="search-result">
          <strong>${escapeHtml(item.title || "untitled")}</strong>
          <span>${escapeHtml(item.snippet || "설명 없음")}</span>
          <a href="${escapeHtml(item.url || "#")}" target="_blank" rel="noreferrer noopener">${escapeHtml(item.url || "")}</a>
          <div class="button-row">
            <button class="secondary search-open" data-url="${encodeURIComponent(item.url || "")}">열기</button>
          </div>
        </div>
      `,
    )
    .join("");
  elements.searchResults.querySelectorAll(".search-open").forEach((button) => {
    button.addEventListener("click", async () => {
      const url = decodeURIComponent(button.dataset.url || "");
      if (url) {
        await runSystemAction("open_url", { url }, `링크 열기: ${url}`);
      }
    });
  });
}

function formatSessionTime(ts) {
  if (!ts) {
    return "-";
  }
  return new Date(ts * 1000).toLocaleTimeString("ko-KR", { hour12: false });
}

function renderSessionEvents(events) {
  if (!elements.sessionLog) {
    return;
  }
  if (!Array.isArray(events) || events.length === 0) {
    elements.sessionLog.textContent = "아직 기록이 없습니다.";
    return;
  }
  elements.sessionLog.innerHTML = [...events]
    .reverse()
    .map((event) => {
      const payload = event.payload || {};
      const summary = payload.prompt || payload.query || payload.target || payload.mode || "";
      return `
        <div class="session-event">
          <strong>${escapeHtml(event.type || "event")}</strong>
          <span>${escapeHtml(formatSessionTime(event.ts))}</span>
          <span>${escapeHtml(summary).slice(0, 180)}</span>
          <code>${escapeHtml(JSON.stringify(payload, null, 2))}</code>
        </div>
      `;
    })
    .join("");
}

async function refreshRuntimeRegistry() {
  try {
    const response = await fetch("/api/runtime", { cache: "no-store" });
    const data = await response.json();
    if (!data.ok) {
      throw new Error("runtime unavailable");
    }
    renderRegistryRows(elements.toolRegistry, data.runtime?.tools, "도구 없음");
    renderRegistryRows(elements.permissionRegistry, data.runtime?.permissions, "권한 정보 없음");
    if (elements.sessionMeta) {
      const runtimeSession = data.runtime?.session || {};
      elements.sessionMeta.textContent = `세션: ${runtimeSession.id || "-"} | 이벤트: ${runtimeSession.eventCount || 0}`;
    }
    if (elements.computerUseMeta) {
      elements.computerUseMeta.textContent = `browser-use reference: ${data.runtime?.computerUse?.reference?.detail || "-"} | sessions ${data.runtime?.computerUse?.sessions ?? 0}`;
    }
  } catch (error) {
    if (elements.toolRegistry) {
      elements.toolRegistry.textContent = "도구 레지스트리 로드 실패";
    }
    if (elements.permissionRegistry) {
      elements.permissionRegistry.textContent = String(error.message || error);
    }
  }
}

async function refreshSessionLog() {
  try {
    const response = await fetch("/api/session", { cache: "no-store" });
    const data = await response.json();
    if (!data.ok) {
      throw new Error("session unavailable");
    }
    if (elements.sessionMeta) {
      elements.sessionMeta.textContent = `세션: ${data.session?.id || "-"} | 최근 이벤트: ${(data.session?.events || []).length}`;
    }
    renderSessionEvents(data.session?.events || []);
  } catch (error) {
    if (elements.sessionLog) {
      elements.sessionLog.textContent = String(error.message || error);
    }
  }
}

async function refreshOnlyOfficeSessions() {
  try {
    const response = await fetch("/api/onlyoffice/sessions", { cache: "no-store" });
    const data = await response.json();
    if (!data.ok) {
      throw new Error("onlyoffice sessions unavailable");
    }
    renderOnlyOfficeSessions(data.sessions || []);
  } catch (error) {
    if (elements.onlyofficeSessions) {
      elements.onlyofficeSessions.textContent = String(error.message || error);
    }
  }
}

async function refreshComputerUseSessions() {
  try {
    const response = await fetch("/api/computer-use/sessions", { cache: "no-store" });
    const data = await response.json();
    if (!data.ok) {
      throw new Error("computer use sessions unavailable");
    }
    renderComputerUseSessions(data.sessions || []);
  } catch (error) {
    if (elements.computerUseSessions) {
      elements.computerUseSessions.textContent = String(error.message || error);
    }
  }
}

function buildGuiPrompt(rawPrompt) {
  const prompt = String(rawPrompt || "").trim();
  const effects = {
    search: Boolean(elements.searchEnabled?.checked),
    modelProfile: elements.modelProfile?.value || "balanced",
  };
  let finalPrompt = prompt;
  if (effects.modelProfile === "fast") {
    finalPrompt = `${finalPrompt}\n\n응답은 짧고 빠르게 작성해.`;
  } else if (effects.modelProfile === "deep") {
    finalPrompt = `${finalPrompt}\n\n더 길고 구조적으로 정리해.`;
  }
  return { prompt: finalPrompt.trim(), effects };
}

function detectAgentRoute(prompt) {
  const text = String(prompt || "").trim().toLowerCase();
  if (!text) {
    return "document";
  }
  const browserKeywords = [
    "브라우저",
    "사이트",
    "링크",
    "검색",
    "찾아줘",
    "열어줘",
    "공식 문서",
    "릴리스",
    "release",
    "latest",
    "open",
    "find",
    "docs",
    "documentation",
  ];
  const documentKeywords = [
    "작성",
    "초안",
    "정리",
    "회의록",
    "보고서",
    "연구노트",
    "문서",
    "표로",
    "슬라이드",
    "다듬어",
  ];
  const browserScore = browserKeywords.filter((keyword) => text.includes(keyword)).length;
  const documentScore = documentKeywords.filter((keyword) => text.includes(keyword)).length;
  if (browserScore > 0 && browserScore >= documentScore) {
    return "computer_use";
  }
  return "document";
}

async function runWebSearch(query) {
  const input = String(query || "").trim();
  if (!input) {
    renderSearchResults([]);
    return;
  }
  if (elements.searchResults) {
    elements.searchResults.textContent = "검색 중...";
  }
  try {
    const response = await fetch("/api/search", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ query: input }),
    });
    const result = await response.json();
    if (!result.ok) {
      throw new Error(result.detail || result.error || "search failed");
    }
    renderSearchResults(result.results || []);
    setStatus("웹 검색 완료", `${(result.results || []).length}건 결과`);
    await refreshSessionLog();
  } catch (error) {
    if (elements.searchResults) {
      elements.searchResults.textContent = String(error.message || error);
    }
    setStatus("웹 검색 실패", String(error.message || error));
  } finally {
  }
}

async function runSystemAction(action, payload, successMessage = "시스템 액션 실행 완료") {
  if (elements.systemActionLog) {
    elements.systemActionLog.textContent = "시스템 액션 실행 중...";
  }
  try {
    const response = await fetch("/api/system-action", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action, ...payload }),
    });
    const result = await response.json();
    if (!result.ok) {
      throw new Error(result.detail || result.error || "system action failed");
    }
    if (elements.systemActionLog) {
      elements.systemActionLog.textContent = `${successMessage}\n${JSON.stringify(result.result, null, 2)}`;
    }
    await refreshSessionLog();
    return result.result;
  } catch (error) {
    if (elements.systemActionLog) {
      elements.systemActionLog.textContent = String(error.message || error);
    }
    throw error;
  }
}

async function planComputerUse(goal, options = {}) {
  const input = String(goal || elements.promptInput?.value || "").trim();
  if (!input) {
    if (elements.computerUseMeta) {
      elements.computerUseMeta.textContent = "브라우저 목표를 입력해야 합니다.";
    }
    return;
  }
  setLiveRoute("computer_use", input);
  if (elements.computerUseMeta) {
    elements.computerUseMeta.textContent = "브라우저 작업 계획 생성 중...";
  }
  setBadge("브라우저 계획");
  try {
    const response = await fetch("/api/computer-use/plan", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ goal: input }),
    });
    const result = await response.json();
    if (!result.ok) {
      throw new Error(result.detail || result.error || "computer use plan failed");
    }
    state.currentComputerUseSessionId = result.session_id;
    state.currentComputerUsePlan = {
      id: result.session_id,
      status: "planned",
      executed_steps: 0,
      goal: input,
      plan: result.plan,
      actions: result.plan.actions,
      summary: result.plan.summary,
    };
    renderComputerUsePlan(state.currentComputerUsePlan);
    if (elements.computerUseMeta) {
      elements.computerUseMeta.textContent = `브라우저 계획 생성 완료 · ${result.plan.meta?.planner || "-"} · ${result.plan.actions.length}단계`;
    }
    rememberCommand(input, "computer_use");
    persistWorkspace();
    setStatus("브라우저 계획 생성 완료", `${result.plan.actions.length}단계 · ${result.plan.summary || input}`);
    setWorkflowHint(`브라우저 작업 계획 생성: ${input}`);
    await refreshComputerUseSessions();
    await refreshSessionLog();
    await refreshRuntimeRegistry();
    await refreshAgentHealth();
    if (options.autorun) {
      await runAllComputerUseSteps();
    }
  } catch (error) {
    if (elements.computerUseMeta) {
      elements.computerUseMeta.textContent = String(error.message || error);
    }
    if (elements.computerUsePlan) {
      elements.computerUsePlan.textContent = String(error.message || error);
    }
  } finally {
  }
}

function getNextComputerUseStepIndex() {
  const session = state.currentComputerUsePlan;
  if (!session) {
    return -1;
  }
  const progress = summarizeComputerUseProgress(session);
  return progress.done < progress.total ? progress.done : -1;
}

async function runComputerUseStep(sessionId, stepIndex) {
  if (!sessionId || stepIndex < 0) {
    return;
  }
  state.computerUseBusy = true;
  setLiveRoute("computer_use");
  setComputerUseProgress(state.currentComputerUsePlan);
  if (elements.computerUseMeta) {
    elements.computerUseMeta.textContent = `브라우저 단계 실행 중... 세션 ${sessionId} / 단계 ${stepIndex + 1}`;
  }
  setBadge(`브라우저 ${stepIndex + 1}단계`);
  setStatus("브라우저 단계 실행 중", `세션 ${sessionId} · 단계 ${stepIndex + 1}`);
  try {
    const response = await fetch("/api/computer-use/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: sessionId, step_index: stepIndex }),
    });
    const result = await response.json();
    if (!result.ok) {
      throw new Error(result.detail || result.error || "computer use run failed");
    }
    const detail = result.result?.action?.label || result.result?.action?.type || "step";
    if (state.currentComputerUsePlan && state.currentComputerUseSessionId === sessionId) {
      state.currentComputerUsePlan.executed_steps = Math.max(Number(state.currentComputerUsePlan.executed_steps || 0), stepIndex + 1);
      state.currentComputerUsePlan.status = result.result?.action?.type === "note" ? "review" : "running";
      renderComputerUsePlan(state.currentComputerUsePlan);
    }
    if (elements.computerUseMeta) {
      elements.computerUseMeta.textContent = `실행 완료: ${detail}`;
    }
    setStatus("브라우저 단계 실행 완료", detail);
    setBadge("브라우저 진행");
    await refreshComputerUseSessions();
    await refreshSessionLog();
    await refreshRuntimeRegistry();
    await refreshAgentHealth();
  } catch (error) {
    if (elements.computerUseMeta) {
      elements.computerUseMeta.textContent = String(error.message || error);
    }
    setStatus("브라우저 단계 실행 실패", String(error.message || error));
    setBadge("실패");
  } finally {
    state.computerUseBusy = false;
    setComputerUseProgress(state.currentComputerUsePlan);
  }
}

async function runNextComputerUseStep() {
  const sessionId = state.currentComputerUseSessionId;
  const stepIndex = getNextComputerUseStepIndex();
  if (!sessionId || stepIndex < 0) {
    if (elements.computerUseMeta) {
      elements.computerUseMeta.textContent = "실행할 다음 단계가 없습니다.";
    }
    return;
  }
  await runComputerUseStep(sessionId, stepIndex);
}

async function runAllComputerUseSteps() {
  if (state.computerUseBusy) {
    return;
  }
  setLiveRoute("computer_use");
  setBadge("자동 진행");
  while (true) {
    const sessionId = state.currentComputerUseSessionId;
    const stepIndex = getNextComputerUseStepIndex();
    if (!sessionId || stepIndex < 0) {
      break;
    }
    await runComputerUseStep(sessionId, stepIndex);
    await new Promise((resolve) => window.setTimeout(resolve, 180));
  }
  if (elements.computerUseMeta) {
    elements.computerUseMeta.textContent = "자동 진행 완료";
  }
  setStatus("브라우저 자동 진행 완료");
  setBadge("완료");
}

function parseJson(value) {
  try {
    return JSON.parse(value);
  } catch {
    return null;
  }
}

function escapeHtml(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function textToSafeHtml(text) {
  return String(text || "")
    .split(/\n{2,}/)
    .map((chunk) => `<p>${escapeHtml(chunk).replaceAll("\n", "<br>")}</p>`)
    .join("");
}

function blocksToText(blocks) {
  if (!Array.isArray(blocks)) {
    return "";
  }
  const lines = [];
  blocks.forEach((block) => {
    if (!block || typeof block !== "object") {
      return;
    }
    if (block.kind === "heading") {
      if (block.text) {
        lines.push(String(block.text).trim());
        lines.push("");
      }
      return;
    }
    if (block.kind === "paragraph") {
      if (block.text) {
        lines.push(String(block.text).trim());
        lines.push("");
      }
      return;
    }
    if (block.kind === "bullets") {
      (block.items || []).forEach((item) => lines.push(`- ${String(item).trim()}`));
      lines.push("");
      return;
    }
    if (block.kind === "numbered") {
      (block.items || []).forEach((item, index) => lines.push(`${index + 1}. ${String(item).trim()}`));
      lines.push("");
      return;
    }
    if (block.kind === "table") {
      const headers = Array.isArray(block.headers) ? block.headers.map((item) => String(item).trim()) : [];
      if (headers.length > 0) {
        lines.push(headers.join(" | "));
      }
      (block.rows || []).forEach((row) => {
        if (Array.isArray(row)) {
          lines.push(row.map((cell) => String(cell).trim()).join(" | "));
        }
      });
      lines.push("");
    }
  });
  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

function htmlToStructuredText(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div>${String(html || "")}</div>`, "text/html");
  const lines = [];

  function walk(node, indent = "") {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = (node.textContent || "").replace(/\s+/g, " ").trim();
      if (text) {
        lines.push(`${indent}${text}`);
      }
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const tag = node.tagName.toUpperCase();
    if (["H1", "H2", "H3", "P"].includes(tag)) {
      const text = (node.textContent || "").replace(/\s+/g, " ").trim();
      if (text) {
        lines.push(text);
        lines.push("");
      }
      return;
    }

    if (tag === "LI") {
      const text = (node.textContent || "").replace(/\s+/g, " ").trim();
      if (text) {
        lines.push(`${indent}- ${text}`);
      }
      return;
    }

    if (tag === "TR") {
      const cells = [...node.children]
        .filter((child) => ["TH", "TD"].includes(child.tagName.toUpperCase()))
        .map((child) => (child.textContent || "").replace(/\s+/g, " ").trim());
      if (cells.length > 0) {
        lines.push(cells.join(" | "));
      }
      return;
    }

    [...node.childNodes].forEach((child) => walk(child, indent));
    if (["UL", "OL", "TABLE", "TBODY", "THEAD"].includes(tag)) {
      lines.push("");
    }
  }

  [...doc.body.firstChild.childNodes].forEach((child) => walk(child));
  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

function normalizeHtmlForRhwp(html) {
  const parser = new DOMParser();
  const source = String(html || "").replace(/<!--[\s\S]*?-->/g, "");
  const doc = parser.parseFromString(`<div>${source}</div>`, "text/html");
  const allowed = new Set(["H1", "H2", "H3", "P", "UL", "OL", "LI", "STRONG", "EM", "BR", "TABLE", "THEAD", "TBODY", "TR", "TH", "TD"]);

  function sanitizeNode(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      return doc.createTextNode(node.textContent || "");
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return doc.createTextNode("");
    }

    const tag = node.tagName.toUpperCase();
    if (!allowed.has(tag)) {
      const fragment = doc.createDocumentFragment();
      [...node.childNodes].forEach((child) => fragment.appendChild(sanitizeNode(child)));
      return fragment;
    }

    const clean = doc.createElement(tag.toLowerCase());
    [...node.childNodes].forEach((child) => clean.appendChild(sanitizeNode(child)));
    return clean;
  }

  const container = doc.createElement("div");
  [...doc.body.firstChild.childNodes].forEach((child) => container.appendChild(sanitizeNode(child)));
  const normalized = container.innerHTML.trim();
  return normalized || "<p>내용 없음</p>";
}

function safeParagraphCount() {
  try {
    return state.doc.getParagraphCount(0);
  } catch {
    return 1;
  }
}

function clearWriterDocument() {
  createBlankDocument();
}

function insertParagraphText(paraIndex, text) {
  if (!text) {
    return;
  }
  state.doc.insertText(0, paraIndex, 0, text);
}

function appendParagraphAfter(paraIndex) {
  const length = getParagraphLength(0, paraIndex);
  state.doc.splitParagraph(0, paraIndex, length);
  return paraIndex + 1;
}

function replaceWriterWithText(text) {
  clearWriterDocument();
  const paragraphs = String(text || "")
    .split(/\n{2,}/)
    .map((item) => item.trim())
    .filter(Boolean);

  if (paragraphs.length === 0) {
    return;
  }

  let paraIndex = 0;
  paragraphs.forEach((paragraphText, index) => {
    insertParagraphText(paraIndex, paragraphText);
    if (index < paragraphs.length - 1) {
      paraIndex = appendParagraphAfter(paraIndex);
    }
  });
}

function rebuildWriterFromParagraphItems(items) {
  const paragraphs = (items || []).map((item) => String(item || "").trim()).filter(Boolean);
  replaceWriterWithText(paragraphs.join("\n\n"));
  persistWorkspace();
}

function getWriterEditorValues() {
  if (!elements.writerEditor) {
    return [];
  }
  return [...elements.writerEditor.querySelectorAll(".writer-paragraph-input")].map((node) => node.value);
}

function moveListItem(items, fromIndex, toIndex) {
  if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0 || fromIndex >= items.length || toIndex >= items.length) {
    return items;
  }
  const next = [...items];
  const [item] = next.splice(fromIndex, 1);
  next.splice(toIndex, 0, item);
  return next;
}

function syncWriterFromEditor(options = {}) {
  const { rerenderEditor = false, status = "" } = options;
  rebuildWriterFromParagraphItems(getWriterEditorValues());
  const summary = getDocumentSummary();
  updateMeta(summary);
  renderPages();
  if (rerenderEditor) {
    renderWriterEditor(summary);
  }
  if (status) {
    setStatus(status);
  }
}

function appendWriterText(text) {
  const paragraphs = String(text || "")
    .split(/\n{2,}/)
    .map((item) => item.trim())
    .filter(Boolean);
  if (paragraphs.length === 0) {
    return;
  }
  let paraIndex = Math.max(0, safeParagraphCount() - 1);
  let offset = getParagraphLength(0, paraIndex);
  if (offset > 0) {
    paraIndex = appendParagraphAfter(paraIndex);
  }
  paragraphs.forEach((paragraphText, index) => {
    insertParagraphText(paraIndex, paragraphText);
    if (index < paragraphs.length - 1) {
      paraIndex = appendParagraphAfter(paraIndex);
    }
  });
}

function insertWriterTemplate(kind) {
  if (!state.doc) {
    createBlankDocument();
  }
  if (kind === "heading") {
    appendWriterText("제목");
    persistWorkspace();
    return;
  }
  if (kind === "paragraph") {
    appendWriterText("새 문단 내용을 입력하세요.");
    persistWorkspace();
    return;
  }
  if (kind === "table") {
    appendWriterText("항목 | 내용 | 비고\n1 | 내용 입력 | -\n2 | 내용 입력 | -");
    persistWorkspace();
  }
}

function replaceParagraphWithText(section, paragraph, text) {
  if (!paragraphExists(section, paragraph)) {
    throw new Error(`문단 없음: ${section}:${paragraph}`);
  }
  const length = getParagraphLength(section, paragraph);
  if (length > 0) {
    state.doc.deleteText(section, paragraph, 0, length);
  }
  state.doc.insertText(section, paragraph, 0, text);
}

function applyWriterHtml(section, paragraph, offset, html, mode = "replace_all") {
  const safeHtml = normalizeHtmlForRhwp(html);
  const plainText = htmlToStructuredText(safeHtml);
  if (mode === "replace_all") {
    replaceWriterWithText(plainText);
    return;
  }
  if (mode === "append") {
    appendWriterText(plainText);
    return;
  }
  replaceParagraphWithText(section, paragraph, plainText);
}

function applyWriterBlocks(section, paragraph, blocks, mode = "replace_all") {
  const plainText = blocksToText(blocks);
  if (mode === "replace_all") {
    replaceWriterWithText(plainText);
    return;
  }
  if (mode === "append") {
    appendWriterText(plainText);
    return;
  }
  replaceParagraphWithText(section, paragraph, plainText);
}

function createBlankDocument() {
  if (state.doc) {
    state.doc.free();
  }
  const doc = HwpDocument.createEmpty();
  doc.createBlankDocument();
  state.doc = doc;
  state.fileName = "untitled.hwp";
  persistWorkspace();
}

function downloadText(text, fileName, mimeType = "text/plain;charset=utf-8") {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function xmlText(node, selector) {
  const target = node.querySelector(selector);
  return target?.textContent?.trim() || "";
}

function escapeXml(text) {
  return String(text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function parseXml(xmlTextValue) {
  return new DOMParser().parseFromString(xmlTextValue, "application/xml");
}

async function unzipOfficeFile(file) {
  const bytes = await file.arrayBuffer();
  return JSZip.loadAsync(bytes);
}

async function parseDocxFile(file) {
  const zip = await unzipOfficeFile(file);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("DOCX 문서 본문을 찾지 못했습니다.");
  }
  const xml = parseXml(documentXml);
  const paragraphs = [...xml.querySelectorAll("w\\:body > w\\:p, body > p")]
    .map((paragraph) =>
      [...paragraph.querySelectorAll("w\\:t, t")]
        .map((node) => node.textContent || "")
        .join("")
        .trim(),
    )
    .filter(Boolean);

  if (paragraphs.length === 0) {
    throw new Error("DOCX에서 읽을 문단이 없습니다.");
  }

  createBlankDocument();
  replaceWriterWithText(paragraphs.join("\n\n"));
  state.fileName = file.name.replace(/\.docx$/i, ".hwp");
  setMode("writer");
  persistWorkspace();
  setStatus("DOCX 문서를 Writer로 가져왔습니다.", `${paragraphs.length}개 문단을 불러왔습니다.`);
}

async function parseXlsxFile(file) {
  const zip = await unzipOfficeFile(file);
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!workbookXml) {
    throw new Error("XLSX 워크북을 찾지 못했습니다.");
  }

  const sharedStringsXml = await zip.file("xl/sharedStrings.xml")?.async("string");
  const sharedStrings = sharedStringsXml
    ? [...parseXml(sharedStringsXml).querySelectorAll("si")]
        .map((item) => [...item.querySelectorAll("t")].map((node) => node.textContent || "").join("").trim())
    : [];

  const workbook = parseXml(workbookXml);
  const firstSheet = workbook.querySelector("sheet");
  const relationshipId = firstSheet?.getAttribute("r:id");
  if (!relationshipId) {
    throw new Error("XLSX 첫 시트를 찾지 못했습니다.");
  }

  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  const rels = relsXml ? parseXml(relsXml) : null;
  const rel = rels?.querySelector(`Relationship[Id="${relationshipId}"]`);
  const target = rel?.getAttribute("Target");
  if (!target) {
    throw new Error("XLSX 시트 경로를 찾지 못했습니다.");
  }

  const normalizedTarget = target.startsWith("/")
    ? target.replace(/^\//, "")
    : target.startsWith("worksheets/")
      ? `xl/${target}`
      : `xl/${target.replace(/^xl\//, "")}`;
  const sheetXml = await zip.file(normalizedTarget)?.async("string");
  if (!sheetXml) {
    throw new Error("XLSX 시트 데이터를 읽지 못했습니다.");
  }

  const rows = [...parseXml(sheetXml).querySelectorAll("sheetData > row")].map((row) =>
    [...row.querySelectorAll("c")].map((cell) => {
      const type = cell.getAttribute("t");
      const value = xmlText(cell, "v");
      if (type === "s") {
        return sharedStrings[Number(value)] || "";
      }
      return value;
    }),
  );

  const headerRow = rows[0]?.length ? rows[0] : SHEET_COLUMNS;
  const columns = headerRow.map((value, index) => value || `열${index + 1}`);
  state.sheetColumns = columns;
  state.sheetRows = rows.slice(1).map((row) =>
    columns.reduce((acc, column, index) => {
      acc[column] = row[index] || "";
      return acc;
    }, {}),
  );
  if (state.sheetRows.length === 0) {
    state.sheetRows = Array.from({ length: 8 }, () =>
      columns.reduce((acc, column) => {
        acc[column] = "";
        return acc;
      }, {}),
    );
  }
  state.fileName = file.name;
  renderSheet();
  setMode("sheet");
  persistWorkspace();
  setStatus("XLSX 시트를 Sheet로 가져왔습니다.", `${state.sheetColumns.length}개 열, ${state.sheetRows.length}개 행을 불러왔습니다.`);
}

async function parsePptxFile(file) {
  const zip = await unzipOfficeFile(file);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const aNumber = Number(a.match(/slide(\d+)\.xml/i)?.[1] || 0);
      const bNumber = Number(b.match(/slide(\d+)\.xml/i)?.[1] || 0);
      return aNumber - bNumber;
    });

  if (slideFiles.length === 0) {
    throw new Error("PPTX 슬라이드를 찾지 못했습니다.");
  }

  state.slides = [];
  for (const [index, slidePath] of slideFiles.entries()) {
    const slideXml = await zip.file(slidePath)?.async("string");
    if (!slideXml) {
      continue;
    }
    const xml = parseXml(slideXml);
    const texts = [...xml.querySelectorAll("a\\:t, t")]
      .map((node) => (node.textContent || "").trim())
      .filter(Boolean);
    if (texts.length === 0) {
      state.slides.push({ title: `슬라이드 ${index + 1}`, bullets: [] });
      continue;
    }
    state.slides.push({
      title: texts[0] || `슬라이드 ${index + 1}`,
      bullets: texts.slice(1, 9),
    });
  }

  renderSlides();
  state.fileName = file.name;
  setMode("slides");
  persistWorkspace();
  setStatus("PPTX 슬라이드를 Slides로 가져왔습니다.", `${state.slides.length}개 슬라이드를 불러왔습니다.`);
}

function toDocxFileName(fileName) {
  return fileName.replace(/\.(hwp|hwpx|docx)$/i, "") + ".docx";
}

function toXlsxFileName(fileName) {
  return fileName.replace(/\.(csv|xlsx)$/i, "") + ".xlsx";
}

function toPptxFileName(fileName) {
  return fileName.replace(/\.(md|pptx)$/i, "") + ".pptx";
}

function toIsoDate(value = new Date()) {
  return value.toISOString().replace(/\.\d{3}Z$/, "Z");
}

async function buildDocxBytes() {
  const summary = getDocumentSummary();
  const paragraphs = summary.paragraphs
    .map((item) => String(item.text || "").trim())
    .filter(Boolean);
  const zip = new JSZip();
  const created = toIsoDate();
  const bodyXml = (paragraphs.length ? paragraphs : [""])
    .map(
      (text) =>
        `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`,
    )
    .join("");
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.folder("_rels")?.file(
    ".rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
  );
  zip.folder("docProps")?.file(
    "app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>hwp</Application>
</Properties>`,
  );
  zip.folder("docProps")?.file(
    "core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(summary.fileName || "hwp Document")}</dc:title>
  <dc:creator>hwp</dc:creator>
  <cp:lastModifiedBy>hwp</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>
</cp:coreProperties>`,
  );
  zip.folder("word")?.file(
    "document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
  <w:body>
    ${bodyXml}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.folder("word")?.folder("_rels")?.file(
    "document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

async function exportDocx() {
  const bytes = await buildDocxBytes();
  downloadBytes(bytes, toDocxFileName(state.fileName || "office-agent-document"));
}

function xlsxColumnName(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

async function buildXlsxBytes() {
  const zip = new JSZip();
  const rows = [
    state.sheetColumns,
    ...state.sheetRows.map((row) => state.sheetColumns.map((column) => String(row[column] ?? ""))),
  ];
  const sheetRowsXml = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((value, columnIndex) => {
          const ref = `${xlsxColumnName(columnIndex)}${rowIndex + 1}`;
          return `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
  );
  zip.folder("_rels")?.file(
    ".rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
  );
  zip.folder("docProps")?.file(
    "app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>hwp</Application>
</Properties>`,
  );
  zip.folder("docProps")?.file(
    "core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(state.fileName || "hwp Sheet")}</dc:title>
  <dc:creator>hwp</dc:creator>
</cp:coreProperties>`,
  );
  zip.folder("xl")?.file(
    "workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  );
  zip.folder("xl")?.folder("_rels")?.file(
    "workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  );
  zip.folder("xl")?.file(
    "styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
  );
  zip.folder("xl")?.folder("worksheets")?.file(
    "sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${sheetRowsXml}</sheetData>
</worksheet>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

async function exportXlsx() {
  const bytes = await buildXlsxBytes();
  downloadBytes(bytes, toXlsxFileName(state.fileName || "office-agent-sheet"));
}

async function buildPptxBytes() {
  const slides = state.slides.length > 0 ? state.slides : [{ title: "슬라이드 1", bullets: [] }];
  const zip = new JSZip();
  const created = toIsoDate();
  const slideEntries = slides
    .map(
      (_slide, index) =>
        `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`,
    )
    .join("");
  const slideRels = slides
    .map(
      (_slide, index) =>
        `<Relationship Id="rId${index + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`,
    )
    .join("");
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  ${slides
    .map(
      (_slide, index) =>
        `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
    )
    .join("\n  ")}
</Types>`,
  );
  zip.folder("_rels")?.file(
    ".rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
  );
  zip.folder("docProps")?.file(
    "app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>hwp</Application>
  <Slides>${slides.length}</Slides>
</Properties>`,
  );
  zip.folder("docProps")?.file(
    "core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(state.fileName || "hwp Slides")}</dc:title>
  <dc:creator>hwp</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>
</cp:coreProperties>`,
  );
  zip.folder("ppt")?.file(
    "presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${slideEntries}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`,
  );
  zip.folder("ppt")?.folder("_rels")?.file(
    "presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
</Relationships>`,
  );
  zip.folder("ppt")?.file(
    "presProps.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
  );
  zip.folder("ppt")?.file(
    "viewProps.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
  );
  zip.folder("ppt")?.file(
    "tableStyles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
  );
  zip.folder("ppt")?.folder("theme")?.file(
    "theme1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="hwp Theme">
  <a:themeElements>
    <a:clrScheme name="Default">
      <a:dk1><a:srgbClr val="1F1A17"/></a:dk1>
      <a:lt1><a:srgbClr val="FFF8EF"/></a:lt1>
      <a:accent1><a:srgbClr val="C96F3B"/></a:accent1>
      <a:accent2><a:srgbClr val="51463C"/></a:accent2>
      <a:accent3><a:srgbClr val="E6C08B"/></a:accent3>
      <a:accent4><a:srgbClr val="8E7B66"/></a:accent4>
      <a:accent5><a:srgbClr val="6C8A6B"/></a:accent5>
      <a:accent6><a:srgbClr val="496D8D"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Default">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Default"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>`,
  );
  zip.folder("ppt")?.folder("slideMasters")?.file(
    "slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="hwp Master"><p:spTree/></p:cSld>
  <p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt1" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk1"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`,
  );
  zip.folder("ppt")?.folder("slideMasters")?.folder("_rels")?.file(
    "slideMaster1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`,
  );
  zip.folder("ppt")?.folder("slideLayouts")?.file(
    "slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="titleOnly" preserve="1">
  <p:cSld name="Title Only"><p:spTree/></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`,
  );
  zip.folder("ppt")?.folder("slideLayouts")?.folder("_rels")?.file(
    "slideLayout1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`,
  );
  slides.forEach((slide, index) => {
    const bulletsXml = (slide.bullets || [])
      .map(
        (bullet) =>
          `<a:p><a:pPr marL="457200" indent="-228600"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="ko-KR" sz="2200"/><a:t>${escapeXml(bullet)}</a:t></a:r></a:p>`,
      )
      .join("");
    zip.folder("ppt")?.folder("slides")?.file(
      `slide${index + 1}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="457200"/><a:ext cx="7772400" cy="914400"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="2800" b="1"/><a:t>${escapeXml(slide.title || `슬라이드 ${index + 1}`)}</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="1600200"/><a:ext cx="7772400" cy="2743200"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${bulletsXml || "<a:p/>"}</p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`,
    );
    zip.folder("ppt")?.folder("slides")?.folder("_rels")?.file(
      `slide${index + 1}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
}

async function exportPptx() {
  const bytes = await buildPptxBytes();
  downloadBytes(bytes, toPptxFileName(state.fileName || "office-agent-slides"));
}

async function loadDocumentFromFile(file) {
  const name = file.name.toLowerCase();
  const extension = name.includes(".") ? name.split(".").pop() : "";

  if (["txt", "md"].includes(extension)) {
    const text = await file.text();
    state.noteText = text;
    elements.notesPad.value = text;
    state.fileName = file.name;
    setMode("notes");
    persistWorkspace();
    return;
  }

  if (extension === "csv") {
    const text = await file.text();
    const rows = text.trim().split(/\r?\n/).map((line) => line.split(",").map((item) => item.trim()));
    const columns = rows[0] || SHEET_COLUMNS;
    state.sheetRows = rows.slice(1).map((row) =>
      columns.reduce((acc, column, index) => {
        acc[column] = row[index] || "";
        return acc;
      }, {}),
    );
    state.fileName = file.name;
    renderSheet();
    setMode("sheet");
    persistWorkspace();
    return;
  }

  if (extension === "json") {
    const text = await file.text();
    try {
      const parsed = JSON.parse(text);
      if (isWorkspaceSnapshot(parsed)) {
        createBlankDocument();
        await applyWorkspaceSnapshot(parsed);
        renderSheet();
        renderSlides();
        renderCommandHistory();
        previewAgentRoute();
        setMode(state.mode || "writer");
        persistWorkspace();
        setStatus("작업 패키지를 복원했습니다.", file.name);
        return;
      }
      if (Array.isArray(parsed.slides)) {
        state.slides = parsed.slides;
        renderSlides();
        setMode("slides");
      } else {
        state.noteText = JSON.stringify(parsed, null, 2);
        elements.notesPad.value = state.noteText;
        setMode("notes");
      }
    } catch {
      state.noteText = text;
      elements.notesPad.value = text;
      setMode("notes");
    }
    state.fileName = file.name;
    persistWorkspace();
    return;
  }

  if (extension === "docx") {
    await parseDocxFile(file);
    return;
  }

  if (extension === "xlsx") {
    await parseXlsxFile(file);
    return;
  }

  if (extension === "pptx") {
    await parsePptxFile(file);
    return;
  }

  const bytes = new Uint8Array(await file.arrayBuffer());
  if (state.doc) {
    state.doc.free();
  }
  state.doc = new HwpDocument(bytes);
  state.fileName = file.name;
  setMode("writer");
  persistWorkspace();
}

function downloadBytes(bytes, fileName) {
  const blob = new Blob([bytes], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

async function fetchMemoryExport() {
  const response = await fetch("/api/memory/export", { cache: "no-store" });
  const data = await response.json();
  if (!data.ok) {
    throw new Error(data.detail || data.error || "memory export failed");
  }
  return Array.isArray(data.items) ? data.items : [];
}

async function importWorkspaceMemories(items, replace = true) {
  const response = await fetch("/api/memory/import", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ items, replace }),
  });
  const data = await response.json();
  if (!data.ok) {
    throw new Error(data.detail || data.error || "memory import failed");
  }
  return data.items || [];
}

async function exportWorkspacePackage() {
  const snapshot = createWorkspaceSnapshot();
  try {
    snapshot.memory = await fetchMemoryExport();
  } catch {
    snapshot.memory = [];
  }
  downloadText(JSON.stringify(snapshot, null, 2), toWorkspaceFileName(state.fileName), "application/json;charset=utf-8");
}

function bytesToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

async function openOnlyOfficeSession() {
  let extension = "";
  let title = state.fileName || "document";
  let bytes = null;

  if (state.mode === "writer") {
    extension = "docx";
    title = toDocxFileName(title || "hwp-document");
    bytes = await buildDocxBytes();
  } else if (state.mode === "sheet") {
    extension = "xlsx";
    title = toXlsxFileName(title || "hwp-sheet");
    bytes = await buildXlsxBytes();
  } else if (state.mode === "slides") {
    extension = "pptx";
    title = toPptxFileName(title || "hwp-slides");
    bytes = await buildPptxBytes();
  } else {
    throw new Error("ONLYOFFICE는 Writer, Sheet, Slides에서만 사용합니다.");
  }

  const response = await fetch("/api/onlyoffice/config", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      mode: state.mode,
      title,
      extension,
      content_base64: bytesToBase64(bytes),
    }),
  });
  const result = await response.json();
  if (!result.ok) {
    throw new Error(result.detail || result.error || "onlyoffice launch failed");
  }
  window.open(result.launch_url, "_blank", "noopener,noreferrer");
  setEngineMeta(`ONLYOFFICE 세션 준비 완료: ${title}`);
  await refreshSessionLog();
  await refreshOnlyOfficeSessions();
}

function resetSheetData() {
  state.sheetColumns = [...SHEET_COLUMNS];
  state.sheetRows = Array.from({ length: 12 }, (_, rowIndex) =>
    state.sheetColumns.reduce((acc, column, columnIndex) => {
      acc[column] = rowIndex === 0 && columnIndex === 0 ? "예시 업무" : "";
      return acc;
    }, {}),
  );
  persistWorkspace();
}

function createEmptySheetRow() {
  return state.sheetColumns.reduce((acc, column) => {
    acc[column] = "";
    return acc;
  }, {});
}

function addSheetRow() {
  state.sheetRows.push(createEmptySheetRow());
  renderSheet();
  persistWorkspace();
}

function nextSheetColumnName() {
  return `열${state.sheetColumns.length + 1}`;
}

function addSheetColumn(name = nextSheetColumnName()) {
  const columnName = String(name || nextSheetColumnName()).trim() || nextSheetColumnName();
  state.sheetColumns.push(columnName);
  state.sheetRows = state.sheetRows.map((row) => ({ ...row, [columnName]: "" }));
  renderSheet();
  persistWorkspace();
}

function addSheetTotalsRow() {
  const totals = state.sheetColumns.reduce((acc, column, index) => {
    if (index === 0) {
      acc[column] = "합계";
      return acc;
    }
    const numbers = state.sheetRows
      .map((row) => Number(String(row[column] || "").replace(/,/g, "")))
      .filter((value) => Number.isFinite(value) && value !== 0);
    acc[column] = numbers.length > 0 ? String(numbers.reduce((sum, value) => sum + value, 0)) : "";
    return acc;
  }, {});
  state.sheetRows.push(totals);
  renderSheet();
  persistWorkspace();
}

function renderSheet() {
  const table = document.createElement("table");
  table.className = "sheet-table";
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const indexHead = document.createElement("th");
  indexHead.textContent = "#";
  headerRow.append(indexHead);
  state.sheetColumns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    headerRow.append(th);
  });
  thead.append(headerRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  state.sheetRows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    const indexCell = document.createElement("th");
    indexCell.textContent = String(rowIndex + 1);
    tr.append(indexCell);
    state.sheetColumns.forEach((column) => {
      const td = document.createElement("td");
      td.contentEditable = "true";
      td.textContent = row[column] || "";
      td.addEventListener("input", () => {
        state.sheetRows[rowIndex][column] = td.textContent ?? "";
        persistWorkspace();
      });
      tr.append(td);
    });
    tbody.append(tr);
  });
  table.append(tbody);
  elements.sheetGrid.innerHTML = "";
  elements.sheetGrid.append(table);
}

function exportSheetCsv() {
  const rows = [
    state.sheetColumns.join(","),
    ...state.sheetRows.map((row) =>
      state.sheetColumns.map((column) => `"${String(row[column] || "").replaceAll('"', '""')}"`).join(","),
    ),
  ];
  downloadText(rows.join("\n"), "office-agent-sheet.csv", "text/csv;charset=utf-8");
}

function generateSlidesFromPrompt(prompt) {
  const base = prompt.trim() || "오피스 에이전트 소개";
  state.slides = [
    { title: "제목", bullets: [base, "발표 목적과 배경", "핵심 한 줄 메시지"] },
    { title: "현황", bullets: ["현재 문제점 정리", "업무 흐름 요약", "주요 리스크"] },
    { title: "제안", bullets: ["개선 방안", "실행 방법", "기대 효과"] },
    { title: "실행 계획", bullets: ["단계별 일정", "담당 역할", "필요 자원"] },
    { title: "마무리", bullets: ["결정 필요 사항", "다음 액션", "질의응답"] },
  ];
  renderSlides();
  persistWorkspace();
}

function addSlideCard() {
  state.slides.push({ title: `슬라이드 ${state.slides.length + 1}`, bullets: ["핵심 내용을 입력하세요."] });
  renderSlides();
  persistWorkspace();
}

function renderSlides() {
  elements.slidesDeck.innerHTML = "";
  if (state.slides.length === 0) {
    elements.slidesDeck.innerHTML = "<div class='slide-card'><h4>슬라이드 없음</h4><p>오른쪽 요청 입력 후 생성 버튼을 누르면 발표 초안이 만들어집니다.</p></div>";
    return;
  }
  state.slides.forEach((slide, index) => {
    const card = document.createElement("article");
    card.className = "slide-card";
    const title = document.createElement("h4");
    title.contentEditable = "true";
    title.textContent = slide.title || `슬라이드 ${index + 1}`;
    title.addEventListener("input", () => {
      state.slides[index].title = title.textContent?.trim() || `슬라이드 ${index + 1}`;
      persistWorkspace();
    });
    card.append(title);
    const bulletsBox = document.createElement("div");
    bulletsBox.className = "slide-bullets";
    (slide.bullets || []).forEach((bullet, bulletIndex) => {
      const p = document.createElement("p");
      p.contentEditable = "true";
      p.textContent = bullet;
      p.addEventListener("input", () => {
        state.slides[index].bullets[bulletIndex] = p.textContent?.trim() || "";
        persistWorkspace();
      });
      bulletsBox.append(p);
    });
    const addBulletButton = document.createElement("button");
    addBulletButton.className = "secondary";
    addBulletButton.textContent = "불릿 추가";
    addBulletButton.addEventListener("click", () => {
      state.slides[index].bullets.push("새 불릿");
      renderSlides();
      persistWorkspace();
    });
    card.append(bulletsBox, addBulletButton);
    elements.slidesDeck.append(card);
  });
}

function createWorkspaceSnapshot() {
  return {
    version: 1,
    mode: state.mode,
    fileName: state.fileName,
    writer: getWriterSnapshot(),
    promptInput: elements.promptInput?.value || "",
    liveRoute: state.liveRoute,
    noteText: elements.notesPad.value,
    sheet: {
      columns: state.sheetColumns,
      rows: state.sheetRows,
    },
    slides: state.slides,
    commandHistory: state.commandHistory,
    project: state.project,
    preferences: state.preferences,
  };
}

function toWorkspaceFileName(fileName) {
  const base = String(fileName || "hwp-workspace").replace(/\.(hwp|hwpx|docx|xlsx|pptx|txt|md|csv|json)$/i, "");
  return `${base || "hwp-workspace"}.hwp-workspace.json`;
}

function isWorkspaceSnapshot(value) {
  return Boolean(
    value &&
      typeof value === "object" &&
      (Array.isArray(value.writer?.paragraphs) || typeof value.noteText === "string" || Array.isArray(value.commandHistory)),
  );
}

async function applyWorkspaceSnapshot(saved) {
  if (typeof saved.fileName === "string" && saved.fileName.trim()) {
    state.fileName = saved.fileName;
  }
  if (typeof saved.noteText === "string") {
    state.noteText = saved.noteText;
    elements.notesPad.value = saved.noteText;
  }
  if (Array.isArray(saved.writer?.paragraphs) && saved.writer.paragraphs.length > 0) {
    replaceWriterWithText(saved.writer.paragraphs.join("\n\n"));
  }
  if (saved.sheet?.rows && Array.isArray(saved.sheet.rows)) {
    if (Array.isArray(saved.sheet.columns) && saved.sheet.columns.length > 0) {
      state.sheetColumns = saved.sheet.columns.map((column) => String(column));
    }
    state.sheetRows = saved.sheet.rows;
  }
  if (Array.isArray(saved.slides)) {
    state.slides = saved.slides;
  }
  if (typeof saved.mode === "string") {
    state.mode = saved.mode;
  }
  if (Array.isArray(saved.commandHistory)) {
    state.commandHistory = saved.commandHistory;
  }
  if (typeof saved.promptInput === "string" && elements.promptInput) {
    elements.promptInput.value = saved.promptInput;
  }
  if (typeof saved.liveRoute === "string") {
    state.liveRoute = saved.liveRoute;
  }
  if (saved.project && typeof saved.project === "object") {
    state.project = {
      name: String(saved.project.name || "Untitled Project"),
      goal: String(saved.project.goal || ""),
      focusMode: Boolean(saved.project.focusMode),
    };
    updateProjectUi();
  }
  if (saved.preferences && typeof saved.preferences === "object") {
    state.preferences = {
      onboardingDismissed: Boolean(saved.preferences.onboardingDismissed),
      defaultSearchEnabled: Boolean(saved.preferences.defaultSearchEnabled),
      defaultModelProfile: String(saved.preferences.defaultModelProfile || "balanced"),
    };
  }
  if (Array.isArray(saved.memory)) {
    await importWorkspaceMemories(saved.memory, true);
    await refreshMemoryRecall(String(elements.promptInput?.value || "").trim());
  }
}

function persistWorkspace() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(createWorkspaceSnapshot()));
    state.autosavedAt = Date.now();
    if (elements.topbarStatus) {
      elements.topbarStatus.textContent = `자동 저장 ${formatTimestamp(state.autosavedAt)}`;
    }
  } catch {}
}

async function restoreWorkspace() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return;
    }
    const saved = JSON.parse(raw);
    await applyWorkspaceSnapshot(saved);
  } catch {}
}

function exportSlidesMarkdown() {
  const markdown = state.slides
    .map((slide) => `## ${slide.title}\n${slide.bullets.map((bullet) => `- ${bullet}`).join("\n")}`)
    .join("\n\n");
  downloadText(markdown || "## 슬라이드 초안 없음\n", "office-agent-slides.md", "text/markdown;charset=utf-8");
}

function getDocumentSummary() {
  if (!state.doc) {
    return { pageCount: 0, sectionCount: 0, paragraphs: [] };
  }

  const info = parseJson(state.doc.getDocumentInfo()) ?? {};
  const pageCount = state.doc.pageCount();
  const paragraphMap = new Map();
  const seenRuns = new Set();

  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const layout = parseJson(state.doc.getPageTextLayout(pageIndex));
    for (const run of layout?.runs ?? []) {
      if (typeof run.secIdx !== "number" || typeof run.paraIdx !== "number" || typeof run.charStart !== "number") {
        continue;
      }
      if (run.parentParaIdx !== undefined) {
        continue;
      }
      const runKey = `${run.secIdx}:${run.paraIdx}:${run.charStart}:${run.text}`;
      if (seenRuns.has(runKey)) {
        continue;
      }
      seenRuns.add(runKey);
      const paraKey = `${run.secIdx}:${run.paraIdx}`;
      if (!paragraphMap.has(paraKey)) {
        paragraphMap.set(paraKey, []);
      }
      paragraphMap.get(paraKey).push({
        start: run.charStart,
        text: run.text,
      });
    }
  }

  const paragraphs = [...paragraphMap.entries()]
    .map(([key, runs]) => {
      const [section, paragraph] = key.split(":").map(Number);
      const sortedRuns = runs.sort((a, b) => a.start - b.start);
      let text = "";
      for (const run of sortedRuns) {
        if (run.start > text.length) {
          text += " ".repeat(run.start - text.length);
        }
        const prefix = text.slice(0, run.start);
        const suffixStart = run.start + run.text.length;
        const suffix = text.length > suffixStart ? text.slice(suffixStart) : "";
        text = `${prefix}${run.text}${suffix}`;
      }
      return {
        section,
        paragraph,
        text: text.trim(),
      };
    })
    .sort((a, b) => a.section - b.section || a.paragraph - b.paragraph);

  return {
    fileName: state.fileName,
    pageCount,
    sectionCount: info.sectionCount ?? state.doc.getSectionCount(),
    paragraphs,
  };
}

function updateMeta(document) {
  const info = parseJson(state.doc.getDocumentInfo()) ?? {};
  const lines = [
    `파일명: ${state.fileName}`,
    `페이지: ${document.pageCount}`,
    `섹션: ${document.sectionCount}`,
    `문단: ${document.paragraphs.length}`,
    `버전: ${info.version ?? "-"}`,
  ];
  elements.docMeta.textContent = lines.join("\n");

  if (document.paragraphs.length === 0) {
    elements.outlineBox.textContent = "문서가 비어 있습니다.";
    return;
  }
  elements.outlineBox.textContent = document.paragraphs
    .slice(0, 24)
    .map((item) => `[${item.section}:${item.paragraph}] ${item.text || "(빈 문단)"}`)
    .join("\n");
}

function renderPages() {
  if (!state.doc) {
    return;
  }
  const pageCount = state.doc.pageCount();
  elements.pages.innerHTML = "";
  for (let i = 0; i < pageCount; i += 1) {
    const wrapper = document.createElement("section");
    wrapper.className = "page";
    const label = document.createElement("div");
    label.className = "page-label";
    label.textContent = `Page ${i + 1}`;
    const body = document.createElement("div");
    body.innerHTML = state.doc.renderPageSvg(i);
    wrapper.append(label, body);
    elements.pages.append(wrapper);
  }
}

function renderWriterEditor(document) {
  if (!elements.writerEditor) {
    return;
  }
  const paragraphs = document?.paragraphs || [];
  if (paragraphs.length === 0) {
    elements.writerEditor.innerHTML = "<div class='writer-paragraph-card'><strong>문단 없음</strong><span>문단 추가로 시작할 수 있습니다.</span></div>";
    return;
  }
  elements.writerEditor.innerHTML = paragraphs
    .map(
      (item, index) => `
        <div class="writer-paragraph-card">
          <div class="writer-paragraph-head">
            <strong>문단 ${index + 1}</strong>
            <div class="writer-paragraph-actions">
              <button class="secondary writer-move-up" data-index="${index}">위로</button>
              <button class="secondary writer-move-down" data-index="${index}">아래로</button>
              <button class="secondary writer-duplicate-paragraph" data-index="${index}">복제</button>
              <button class="secondary writer-delete-paragraph" data-index="${index}">삭제</button>
            </div>
          </div>
          <textarea class="writer-paragraph-input" data-index="${index}">${escapeHtml(item.text || "")}</textarea>
        </div>
      `,
    )
    .join("");
  elements.writerEditor.querySelectorAll(".writer-paragraph-input").forEach((input) => {
    input.addEventListener("input", () => {
      window.clearTimeout(writerEditorSyncTimer);
      writerEditorSyncTimer = window.setTimeout(() => {
        syncWriterFromEditor();
      }, 240);
    });
    input.addEventListener("change", () => {
      window.clearTimeout(writerEditorSyncTimer);
      syncWriterFromEditor();
    });
  });
  elements.writerEditor.querySelectorAll(".writer-delete-paragraph").forEach((button) => {
    button.addEventListener("click", async () => {
      const targetIndex = Number(button.dataset.index || "-1");
      const values = getWriterEditorValues()
        .filter((_, index) => index !== targetIndex);
      rebuildWriterFromParagraphItems(values);
      await refreshDocumentView();
      setStatus("문단을 삭제했습니다.");
    });
  });
  elements.writerEditor.querySelectorAll(".writer-duplicate-paragraph").forEach((button) => {
    button.addEventListener("click", async () => {
      const targetIndex = Number(button.dataset.index || "-1");
      const values = getWriterEditorValues();
      if (targetIndex < 0 || targetIndex >= values.length) {
        return;
      }
      values.splice(targetIndex + 1, 0, values[targetIndex]);
      rebuildWriterFromParagraphItems(values);
      await refreshDocumentView();
      setStatus("문단을 복제했습니다.");
    });
  });
  elements.writerEditor.querySelectorAll(".writer-move-up").forEach((button) => {
    button.addEventListener("click", async () => {
      const targetIndex = Number(button.dataset.index || "-1");
      const values = moveListItem(getWriterEditorValues(), targetIndex, targetIndex - 1);
      rebuildWriterFromParagraphItems(values);
      await refreshDocumentView();
      setStatus("문단을 위로 이동했습니다.");
    });
  });
  elements.writerEditor.querySelectorAll(".writer-move-down").forEach((button) => {
    button.addEventListener("click", async () => {
      const targetIndex = Number(button.dataset.index || "-1");
      const values = moveListItem(getWriterEditorValues(), targetIndex, targetIndex + 1);
      rebuildWriterFromParagraphItems(values);
      await refreshDocumentView();
      setStatus("문단을 아래로 이동했습니다.");
    });
  });
}

function getParagraphLength(section, paragraph) {
  return state.doc.getParagraphLength(section, paragraph);
}

function paragraphExists(section, paragraph) {
  try {
    const count = state.doc.getParagraphCount(section);
    return paragraph >= 0 && paragraph < count;
  } catch {
    return false;
  }
}

function applyOperation(operation) {
  if (!state.doc && [
    "set_document_blocks",
    "set_document_html",
    "append_blocks",
    "append_html",
    "replace_paragraph_text",
    "replace_paragraph_blocks",
    "replace_paragraph_html",
  ].includes(operation.type)) {
    throw new Error("문서가 로드되지 않았습니다.");
  }

  if (operation.type === "set_document_blocks") {
    applyWriterBlocks(0, 0, operation.blocks, "replace_all");
    persistWorkspace();
    return;
  }

  if (operation.type === "set_document_html") {
    applyWriterHtml(0, 0, 0, operation.html, "replace_all");
    persistWorkspace();
    return;
  }

  if (operation.type === "append_blocks") {
    const summary = getDocumentSummary();
    const last = summary.paragraphs.at(-1) ?? { section: 0, paragraph: 0 };
    applyWriterBlocks(last.section, last.paragraph, operation.blocks, "append");
    persistWorkspace();
    return;
  }

  if (operation.type === "append_html") {
    const summary = getDocumentSummary();
    const last = summary.paragraphs.at(-1) ?? { section: 0, paragraph: 0 };
    const offset = getParagraphLength(last.section, last.paragraph);
    applyWriterHtml(last.section, last.paragraph, offset, operation.html, "append");
    persistWorkspace();
    return;
  }

  if (operation.type === "replace_paragraph_text") {
    if (!paragraphExists(operation.section, operation.paragraph)) {
      throw new Error(`문단 없음: ${operation.section}:${operation.paragraph}`);
    }
    const length = getParagraphLength(operation.section, operation.paragraph);
    if (length > 0) {
      state.doc.deleteText(operation.section, operation.paragraph, 0, length);
    }
    if (operation.text) {
      state.doc.insertText(operation.section, operation.paragraph, 0, operation.text);
    }
    return;
  }

  if (operation.type === "replace_paragraph_blocks") {
    if (!paragraphExists(operation.section, operation.paragraph)) {
      throw new Error(`문단 없음: ${operation.section}:${operation.paragraph}`);
    }
    const length = getParagraphLength(operation.section, operation.paragraph);
    if (length > 0) {
      state.doc.deleteText(operation.section, operation.paragraph, 0, length);
    }
    applyWriterBlocks(operation.section, operation.paragraph, operation.blocks, "replace_paragraph");
    persistWorkspace();
    return;
  }

  if (operation.type === "replace_paragraph_html") {
    if (!paragraphExists(operation.section, operation.paragraph)) {
      throw new Error(`문단 없음: ${operation.section}:${operation.paragraph}`);
    }
    const length = getParagraphLength(operation.section, operation.paragraph);
    if (length > 0) {
      state.doc.deleteText(operation.section, operation.paragraph, 0, length);
    }
    applyWriterHtml(operation.section, operation.paragraph, 0, operation.html, "replace_paragraph");
    persistWorkspace();
    return;
  }

  if (operation.type === "set_note_text") {
    state.noteText = operation.text ?? "";
    elements.notesPad.value = state.noteText;
    persistWorkspace();
    return;
  }

  if (operation.type === "set_sheet_data") {
    if (Array.isArray(operation.rows) && operation.rows.length > 0) {
      state.sheetColumns = Array.isArray(operation.columns) && operation.columns.length > 0
        ? operation.columns.map((column) => String(column))
        : [...SHEET_COLUMNS];
      state.sheetRows = operation.rows.map((row) => {
        const normalized = {};
        state.sheetColumns.forEach((column) => {
          normalized[column] = row[column] ?? "";
        });
        return normalized;
      });
      renderSheet();
    } else if (Array.isArray(operation.columns) && operation.columns.length > 0) {
      state.sheetColumns = operation.columns.map((column) => String(column));
      state.sheetRows = Array.from({ length: 8 }, () =>
        state.sheetColumns.reduce((acc, column) => {
          acc[column] = "";
          return acc;
        }, {}),
      );
      renderSheet();
    }
    persistWorkspace();
    return;
  }

  if (operation.type === "set_slides") {
    state.slides = Array.isArray(operation.slides) ? operation.slides : [];
    renderSlides();
    persistWorkspace();
    return;
  }
}

async function refreshDocumentView() {
  const summary = getDocumentSummary();
  updateMeta(summary);
  renderPages();
  renderWriterEditor(summary);
}

async function runAgent() {
  const parsed = buildGuiPrompt(elements.promptInput.value);
  const prompt = parsed.prompt;
  if (!prompt) {
    elements.reply.innerHTML = "<p class='error'>요청 문장을 입력해야 합니다.</p>";
    return;
  }
  const route = detectAgentRoute(prompt);
  setLiveRoute(route, prompt);
  if (route === "computer_use") {
    await planComputerUse(prompt, { autorun: true });
    elements.reply.innerHTML = `<p>${escapeHtml("브라우저 작업으로 분기해 자동 진행을 시작했습니다.")}</p>`;
    return;
  }
  rememberCommand(prompt, "document");
  if (parsed.effects.search) {
    await runWebSearch(prompt);
  }

  if (!state.doc) {
    createBlankDocument();
  }

  setBadge("계획 생성 중");
  setLiveRoute("document", prompt);
  setStatus("에이전트가 문서 스냅샷을 읽고 작업 계획을 생성하는 중입니다.");
  elements.runAgent.disabled = true;

  try {
    const payload = {
      prompt,
      document: getDocumentSummary(),
      mode: state.mode,
      noteText: elements.notesPad.value,
      sheet: {
        columns: state.sheetColumns,
        rows: state.sheetRows,
      },
      slides: state.slides,
    };
    const response = await fetch("/api/plan", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const result = await response.json();
    if (!result.ok) {
      throw new Error(result.detail || result.error || "agent request failed");
    }

    const plan = result.plan;
    elements.planBox.textContent = JSON.stringify(plan, null, 2);
    elements.plannerMeta.textContent = `플래너: ${plan.meta?.planner || "-"}${plan.meta?.reason ? ` | 사유: ${plan.meta.reason}` : ""}${plan.meta?.search_results ? ` | 검색 결과: ${plan.meta.search_results}` : ""}`;
    if (elements.dashboardPlan) {
      elements.dashboardPlan.textContent = `${plan.reply || "작업 완료"} · 작업 ${plan.operations.filter((item) => item.type !== "no_op").length}개`;
    }

    for (const operation of plan.operations) {
      if (operation.type === "no_op") {
        continue;
      }
      applyOperation(operation);
    }

    await refreshDocumentView();
    if (state.mode === "slides" && state.slides.length === 0) {
      generateSlidesFromPrompt(prompt);
    }
    if (state.mode === "notes") {
      setMode("notes");
    } else if (state.mode === "sheet") {
      setMode("sheet");
    } else if (state.mode === "slides") {
      setMode("slides");
    }
    elements.reply.innerHTML = `<p>${escapeHtml(plan.reply || "작업을 적용했습니다.")}</p>`;
    setBadge("적용 완료");
    setStatus(
      "에이전트 작업이 적용되었습니다.",
      `실행된 작업 수: ${plan.operations.filter((item) => item.type !== "no_op").length}`,
    );
    await refreshSessionLog();
  } catch (error) {
    elements.reply.innerHTML = `<p class="error">${escapeHtml(String(error.message || error))}</p>`;
    setBadge("실패");
    setStatus("에이전트 실행 실패", String(error.message || error));
  } finally {
    elements.runAgent.disabled = false;
  }
}

async function boot() {
  installMeasureTextWidth();
  setStatus("WASM 엔진 초기화 중");
  setBadge("초기화");
  setRuntimeBadge("Booting");
  await init({ module_or_path: WASM_URL });
  state.ready = true;
  createBlankDocument();
  await restoreWorkspace();
  renderCommandHistory();
  if (state.sheetRows.length === 0) {
    resetSheetData();
  }
  renderSheet();
  renderSlides();
  await refreshDocumentView();
  await refreshAgentHealth();
  await refreshRuntimeRegistry();
  await refreshSessionLog();
  await refreshMemoryRecall("");
  await refreshOnlyOfficeSessions();
  await refreshComputerUseSessions();
  setMode(state.mode || "writer");
  setStatus(
    "준비 완료",
    "HWP/HWPX 파일을 열거나 오른쪽 에이전트에 업무 문서 요청을 입력하면 됩니다.",
  );
  setBadge("준비 완료");
  setRuntimeBadge("Command Ready");
  setWorkflowHint("문서 작업 버튼, 검색 포함 토글, 모델 프로필을 조합해 실행합니다.");
  previewAgentRoute();
  updateProjectUi();
  syncPreferencesUi();
  if (!state.preferences.onboardingDismissed) {
    openAppModal();
  }
  await applyStartupCommandFromUrl();
  if (elements.editorEngine) {
    elements.editorEngine.value = "native";
  }
}

function applyGuiAction(kind) {
  if (kind === "research_note") {
    setMode("writer");
    elements.promptInput.value = "연구노트 형식으로 정리하고 참고 링크를 포함해줘.";
    setWorkflowHint("연구노트 초안 모드입니다. 검색 포함을 켜고 실행하면 참고 링크까지 정리합니다.");
    return;
  }
  if (kind === "research_complete") {
    setMode("writer");
    if (elements.searchEnabled) {
      elements.searchEnabled.checked = true;
    }
    if (elements.modelProfile) {
      elements.modelProfile.value = "deep";
    }
    elements.promptInput.value =
      "연구노트를 완성형으로 작성해줘. 제목, 배경, 핵심 질문, 조사 요약, 참고 링크, 활용 아이디어, 다음 액션을 포함하고 문단 구조를 분명하게 나눠줘.";
    setWorkflowHint("완성형 연구노트 모드입니다. 검색 포함 + 깊은 작성 프로필로 긴 구조 문서를 만듭니다.");
    return;
  }
  if (kind === "minutes") {
    setMode("writer");
    elements.promptInput.value = "회의록 형태로 정리하고 결정사항과 액션아이템을 표로 넣어줘.";
    setWorkflowHint("회의록 초안 모드입니다. 기존 문서를 회의록 형식으로 정리합니다.");
    return;
  }
  if (kind === "report") {
    setMode("writer");
    elements.promptInput.value = "보고서 형식으로 요약, 핵심 내용, 일정 표를 포함해 작성해줘.";
    setWorkflowHint("보고서 초안 모드입니다. 요약과 일정 중심으로 정리합니다.");
    return;
  }
  if (kind === "slides") {
    setMode("slides");
    elements.promptInput.value = "현재 내용을 발표용 슬라이드 초안으로 변환해줘.";
    setWorkflowHint("슬라이드화 모드입니다. 현재 요청을 발표 구조로 변환합니다.");
  }
}

elements.tabWriter.addEventListener("click", () => setMode("writer"));
elements.tabNotes.addEventListener("click", () => setMode("notes"));
elements.tabSheet.addEventListener("click", () => setMode("sheet"));
elements.tabSlides.addEventListener("click", () => setMode("slides"));
elements.topTabWriter?.addEventListener("click", () => setMode("writer"));
elements.topTabNotes?.addEventListener("click", () => setMode("notes"));
elements.topTabSheet?.addEventListener("click", () => setMode("sheet"));
elements.topTabSlides?.addEventListener("click", () => setMode("slides"));

elements.newDocument.addEventListener("click", async () => {
  createBlankDocument();
  await refreshDocumentView();
  setStatus("빈 문서를 새로 만들었습니다.");
  setBadge("새 문서");
});

elements.exportWorkspace?.addEventListener("click", async () => {
  await exportWorkspacePackage();
  setStatus("작업 패키지를 저장했습니다.", toWorkspaceFileName(state.fileName));
});

elements.insertHeading?.addEventListener("click", async () => {
  insertWriterTemplate("heading");
  await refreshDocumentView();
  setMode("writer");
  setStatus("제목 블록을 추가했습니다.");
});

elements.insertParagraph?.addEventListener("click", async () => {
  insertWriterTemplate("paragraph");
  await refreshDocumentView();
  setMode("writer");
  setStatus("문단 블록을 추가했습니다.");
});

elements.insertTable?.addEventListener("click", async () => {
  insertWriterTemplate("table");
  await refreshDocumentView();
  setMode("writer");
  setStatus("표 템플릿을 추가했습니다.");
});

elements.addWriterParagraph?.addEventListener("click", async () => {
  const items = getDocumentSummary()
    .paragraphs
    .map((item) => item.text);
  items.push("새 문단");
  rebuildWriterFromParagraphItems(items);
  await refreshDocumentView();
  setMode("writer");
  setStatus("문단을 추가했습니다.");
});

elements.newNote.addEventListener("click", () => {
  state.noteText = "";
  elements.notesPad.value = "";
  setMode("notes");
  setStatus("새 메모를 시작했습니다.");
  persistWorkspace();
});

elements.notesPad.addEventListener("input", () => {
  state.noteText = elements.notesPad.value;
  persistWorkspace();
});

elements.exportNote.addEventListener("click", () => {
  downloadText(elements.notesPad.value || "", "office-agent-note.txt");
  setStatus("메모를 TXT로 저장했습니다.");
});

elements.addSheetRow?.addEventListener("click", () => {
  addSheetRow();
  setStatus("시트 행을 추가했습니다.");
});

elements.addSheetColumn?.addEventListener("click", () => {
  addSheetColumn();
  setStatus("시트 열을 추가했습니다.");
});

elements.addSheetTotals?.addEventListener("click", () => {
  addSheetTotalsRow();
  setStatus("합계 행을 추가했습니다.");
});

elements.resetSheet.addEventListener("click", () => {
  resetSheetData();
  renderSheet();
  setMode("sheet");
  setStatus("시트를 초기화했습니다.");
  persistWorkspace();
});

elements.exportSheet.addEventListener("click", () => {
  exportSheetCsv();
  setStatus("시트를 CSV로 저장했습니다.");
});

elements.exportXlsx.addEventListener("click", async () => {
  await exportXlsx();
  setStatus("시트를 XLSX로 저장했습니다.");
});

elements.generateSlides.addEventListener("click", () => {
  generateSlidesFromPrompt(elements.promptInput.value);
  setMode("slides");
  setStatus("현재 요청을 기준으로 슬라이드 초안을 생성했습니다.");
});

elements.addSlide?.addEventListener("click", () => {
  addSlideCard();
  setMode("slides");
  setStatus("슬라이드를 추가했습니다.");
});

elements.exportSlides.addEventListener("click", () => {
  exportSlidesMarkdown();
  setStatus("슬라이드 초안을 Markdown으로 저장했습니다.");
});

elements.exportPptx.addEventListener("click", async () => {
  await exportPptx();
  setStatus("슬라이드를 PPTX로 저장했습니다.");
});

elements.exportDocument.addEventListener("click", () => {
  if (!state.doc) {
    return;
  }
  const bytes = state.doc.exportHwp();
  const fileName = state.fileName.endsWith(".hwp") ? state.fileName : "office-agent-document.hwp";
  downloadBytes(bytes, fileName);
  setStatus("문서를 HWP로 저장했습니다.", fileName);
});

elements.exportDocx.addEventListener("click", async () => {
  if (!state.doc) {
    return;
  }
  await exportDocx();
  setStatus("문서를 DOCX로 저장했습니다.");
});

elements.fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }
  setBadge("로딩 중");
  setStatus(`파일 로드 중: ${file.name}`);
  try {
    await loadDocumentFromFile(file);
    await refreshDocumentView();
    setStatus("파일 로드 완료", file.name);
    setBadge("문서 로드");
  } catch (error) {
    setStatus("파일 로드 실패", String(error.message || error));
    setBadge("실패");
  }
});

elements.runAgent.addEventListener("click", runAgent);
elements.saveMemory?.addEventListener("click", async () => {
  const prompt = String(elements.promptInput?.value || "").trim();
  if (!prompt) {
    setStatus("저장할 내용을 먼저 입력해야 합니다.");
    return;
  }
  try {
    const response = await fetch("/api/memory/add", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        kind: "manual",
        title: prompt.slice(0, 80),
        text: prompt,
      }),
    });
    const data = await response.json();
    if (!data.ok) {
      throw new Error(data.detail || data.error || "memory add failed");
    }
    await refreshMemoryRecall(prompt);
    setStatus("현재 입력을 작업 기억에 저장했습니다.");
  } catch (error) {
    setStatus("기억 저장 실패", String(error.message || error));
  }
});
elements.computerUseRunNext?.addEventListener("click", runNextComputerUseStep);
elements.computerUseRunAll?.addEventListener("click", runAllComputerUseSteps);
elements.promptInput?.addEventListener("keydown", (event) => {
  if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
    event.preventDefault();
    runAgent();
  }
});
elements.promptInput?.addEventListener("input", previewAgentRoute);
elements.projectName?.addEventListener("input", () => {
  state.project.name = String(elements.projectName.value || "").trim() || "Untitled Project";
  updateProjectUi();
  persistWorkspace();
});
elements.projectGoal?.addEventListener("input", () => {
  state.project.goal = String(elements.projectGoal.value || "").trim();
  updateProjectUi();
  persistWorkspace();
});
elements.focusToggle?.addEventListener("click", () => setFocusMode());
elements.settingsToggle?.addEventListener("click", () => {
  syncPreferencesUi();
  openAppModal();
});
elements.closeAppModal?.addEventListener("click", () => {
  state.preferences.onboardingDismissed = Boolean(elements.settingsSkipOnboarding?.checked);
  persistWorkspace();
  closeAppModal();
});
elements.settingsModelProfile?.addEventListener("change", () => {
  state.preferences.defaultModelProfile = String(elements.settingsModelProfile.value || "balanced");
  syncPreferencesUi();
  persistWorkspace();
});
elements.settingsSearchDefault?.addEventListener("change", () => {
  state.preferences.defaultSearchEnabled = Boolean(elements.settingsSearchDefault.checked);
  syncPreferencesUi();
  persistWorkspace();
});
elements.settingsSkipOnboarding?.addEventListener("change", () => {
  state.preferences.onboardingDismissed = Boolean(elements.settingsSkipOnboarding.checked);
  persistWorkspace();
});
elements.templateCards.forEach((button) => {
  button.addEventListener("click", () => applyTemplate(button.dataset.template || ""));
});
elements.openBrowser?.addEventListener("click", () =>
  runSystemAction("open_app", { app: "Safari" }, "브라우저를 열었습니다.").catch((error) =>
    setStatus("시스템 액션 실패", String(error.message || error)),
  ),
);
elements.openFinder?.addEventListener("click", () =>
  runSystemAction("open_app", { app: "Finder" }, "Finder를 열었습니다.").catch((error) =>
    setStatus("시스템 액션 실패", String(error.message || error)),
  ),
);
elements.openMonitor?.addEventListener("click", () =>
  runSystemAction("open_app", { app: "Activity Monitor" }, "활동 모니터를 열었습니다.").catch((error) =>
    setStatus("시스템 액션 실패", String(error.message || error)),
  ),
);

elements.promptChips
  .filter((button) => button.dataset.prompt)
  .forEach((button) => {
  button.addEventListener("click", () => {
    elements.promptInput.value = button.dataset.prompt || "";
    previewAgentRoute();
  });
  });

document.querySelectorAll(".computer-preset").forEach((button) => {
  button.addEventListener("click", () => setComputerUsePresetGoal(button.dataset.goal || ""));
});

elements.actionResearchNote?.addEventListener("click", () => applyGuiAction("research_note"));
elements.actionResearchComplete?.addEventListener("click", () => applyGuiAction("research_complete"));
elements.actionMinutes?.addEventListener("click", () => applyGuiAction("minutes"));
elements.actionReport?.addEventListener("click", () => applyGuiAction("report"));
elements.actionSlides?.addEventListener("click", () => applyGuiAction("slides"));
elements.editorEngine?.addEventListener("change", () => {
  setEngineMeta(
    elements.editorEngine.value === "onlyoffice"
      ? "ONLYOFFICE를 선택했습니다. 새 창에서 OOXML 편집 세션을 엽니다."
      : "Native 편집 엔진을 사용합니다. HWP/HWPX와 로컬 워크스페이스를 직접 다룹니다.",
  );
});
elements.openOnlyOffice?.addEventListener("click", async () => {
  try {
    if (elements.editorEngine?.value !== "onlyoffice") {
      elements.editorEngine.value = "onlyoffice";
    }
    await openOnlyOfficeSession();
    setStatus("ONLYOFFICE 세션을 열었습니다.");
  } catch (error) {
    setStatus("ONLYOFFICE 실행 실패", String(error.message || error));
    setEngineMeta(String(error.message || error));
  }
});

boot().catch((error) => {
  setBadge("실패");
  setStatus("초기화 실패", String(error.message || error));
});
