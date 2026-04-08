import init, { HwpDocument } from "https://cdn.jsdelivr.net/npm/@rhwp/core@0.6.1/rhwp.js";

const WASM_URL = "https://cdn.jsdelivr.net/npm/@rhwp/core@0.6.1/rhwp_bg.wasm";

const elements = {
  newDocument: document.querySelector("#new-document"),
  exportDocument: document.querySelector("#export-document"),
  fileInput: document.querySelector("#file-input"),
  docMeta: document.querySelector("#doc-meta"),
  outlineBox: document.querySelector("#outline-box"),
  statusBox: document.querySelector("#status-box"),
  renderBadge: document.querySelector("#render-badge"),
  pages: document.querySelector("#pages"),
  promptInput: document.querySelector("#prompt-input"),
  runAgent: document.querySelector("#run-agent"),
  reply: document.querySelector("#agent-reply"),
  planBox: document.querySelector("#plan-box"),
  plannerMeta: document.querySelector("#planner-meta"),
  promptChips: [...document.querySelectorAll(".prompt-chip")],
};

const state = {
  doc: null,
  fileName: "untitled.hwp",
  ready: false,
};

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
}

function setBadge(message) {
  elements.renderBadge.textContent = message;
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

function createBlankDocument() {
  if (state.doc) {
    state.doc.free();
  }
  const doc = HwpDocument.createEmpty();
  doc.createBlankDocument();
  state.doc = doc;
  state.fileName = "untitled.hwp";
}

async function loadDocumentFromFile(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (state.doc) {
    state.doc.free();
  }
  state.doc = new HwpDocument(bytes);
  state.fileName = file.name;
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
  if (!state.doc) {
    throw new Error("문서가 로드되지 않았습니다.");
  }

  if (operation.type === "set_document_html") {
    createBlankDocument();
    state.doc.pasteHtml(0, 0, 0, operation.html);
    return;
  }

  if (operation.type === "append_html") {
    const summary = getDocumentSummary();
    const last = summary.paragraphs.at(-1) ?? { section: 0, paragraph: 0 };
    const offset = getParagraphLength(last.section, last.paragraph);
    state.doc.pasteHtml(last.section, last.paragraph, offset, operation.html);
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

  if (operation.type === "replace_paragraph_html") {
    if (!paragraphExists(operation.section, operation.paragraph)) {
      throw new Error(`문단 없음: ${operation.section}:${operation.paragraph}`);
    }
    const length = getParagraphLength(operation.section, operation.paragraph);
    if (length > 0) {
      state.doc.deleteText(operation.section, operation.paragraph, 0, length);
    }
    state.doc.pasteHtml(operation.section, operation.paragraph, 0, operation.html);
    return;
  }
}

async function refreshDocumentView() {
  const summary = getDocumentSummary();
  updateMeta(summary);
  renderPages();
}

async function runAgent() {
  const prompt = elements.promptInput.value.trim();
  if (!prompt) {
    elements.reply.innerHTML = "<p class='error'>요청 문장을 입력해야 합니다.</p>";
    return;
  }

  if (!state.doc) {
    createBlankDocument();
  }

  setBadge("계획 생성 중");
  setStatus("에이전트가 문서 스냅샷을 읽고 작업 계획을 생성하는 중입니다.");
  elements.runAgent.disabled = true;

  try {
    const payload = {
      prompt,
      document: getDocumentSummary(),
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
    elements.plannerMeta.textContent = `플래너: ${plan.meta?.planner || "-"}${plan.meta?.reason ? ` | 사유: ${plan.meta.reason}` : ""}`;

    for (const operation of plan.operations) {
      if (operation.type === "no_op") {
        continue;
      }
      applyOperation(operation);
    }

    await refreshDocumentView();
    elements.reply.innerHTML = `<p>${escapeHtml(plan.reply || "작업을 적용했습니다.")}</p>`;
    setBadge("적용 완료");
    setStatus(
      "에이전트 작업이 적용되었습니다.",
      `실행된 작업 수: ${plan.operations.filter((item) => item.type !== "no_op").length}`,
    );
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
  await init({ module_or_path: WASM_URL });
  state.ready = true;
  createBlankDocument();
  await refreshDocumentView();
  setStatus(
    "준비 완료",
    "HWP/HWPX 파일을 열거나 오른쪽 에이전트에 업무 문서 요청을 입력하면 됩니다.",
  );
  setBadge("준비 완료");
}

elements.newDocument.addEventListener("click", async () => {
  createBlankDocument();
  await refreshDocumentView();
  setStatus("빈 문서를 새로 만들었습니다.");
  setBadge("새 문서");
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

elements.promptChips.forEach((button) => {
  button.addEventListener("click", () => {
    elements.promptInput.value = button.dataset.prompt || "";
  });
});

boot().catch((error) => {
  setBadge("실패");
  setStatus("초기화 실패", String(error.message || error));
});
