import init, { HwpDocument } from "https://cdn.jsdelivr.net/npm/@rhwp/core@0.6.1/rhwp.js";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";

const WASM_URL = "https://cdn.jsdelivr.net/npm/@rhwp/core@0.6.1/rhwp_bg.wasm";

const elements = {
  tabWriter: document.querySelector("#tab-writer"),
  tabNotes: document.querySelector("#tab-notes"),
  tabSheet: document.querySelector("#tab-sheet"),
  tabSlides: document.querySelector("#tab-slides"),
  workspaceWriter: document.querySelector("#workspace-writer"),
  workspaceNotes: document.querySelector("#workspace-notes"),
  workspaceSheet: document.querySelector("#workspace-sheet"),
  workspaceSlides: document.querySelector("#workspace-slides"),
  newDocument: document.querySelector("#new-document"),
  exportDocument: document.querySelector("#export-document"),
  fileInput: document.querySelector("#file-input"),
  newNote: document.querySelector("#new-note"),
  exportNote: document.querySelector("#export-note"),
  notesPad: document.querySelector("#notes-pad"),
  resetSheet: document.querySelector("#reset-sheet"),
  exportSheet: document.querySelector("#export-sheet"),
  sheetGrid: document.querySelector("#sheet-grid"),
  generateSlides: document.querySelector("#generate-slides"),
  exportSlides: document.querySelector("#export-slides"),
  slidesDeck: document.querySelector("#slides-deck"),
  docMeta: document.querySelector("#doc-meta"),
  outlineBox: document.querySelector("#outline-box"),
  statusBox: document.querySelector("#status-box"),
  renderBadge: document.querySelector("#render-badge"),
  modeHint: document.querySelector("#mode-hint"),
  pages: document.querySelector("#pages"),
  promptInput: document.querySelector("#prompt-input"),
  runAgent: document.querySelector("#run-agent"),
  reply: document.querySelector("#agent-reply"),
  planBox: document.querySelector("#plan-box"),
  plannerMeta: document.querySelector("#planner-meta"),
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
};

const STORAGE_KEY = "office-agent-staff-state-v1";

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
  [elements.workspaceWriter, elements.workspaceNotes, elements.workspaceSheet, elements.workspaceSlides].forEach((panel) =>
    panel.classList.remove("active"),
  );
  map[mode][0].classList.add("active");
  map[mode][1].classList.add("active");
  elements.modeHint.textContent = `현재 대상: ${mode[0].toUpperCase()}${mode.slice(1)}`;
  persistWorkspace();
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
    title.textContent = `${index + 1}. ${slide.title}`;
    card.append(title);
    slide.bullets.forEach((bullet) => {
      const p = document.createElement("p");
      p.textContent = `• ${bullet}`;
      card.append(p);
    });
    elements.slidesDeck.append(card);
  });
}

function createWorkspaceSnapshot() {
  return {
    mode: state.mode,
    fileName: state.fileName,
    noteText: elements.notesPad.value,
    sheet: {
      columns: state.sheetColumns,
      rows: state.sheetRows,
    },
    slides: state.slides,
  };
}

function persistWorkspace() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(createWorkspaceSnapshot()));
  } catch {}
}

function restoreWorkspace() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return;
    }
    const saved = JSON.parse(raw);
    if (typeof saved.noteText === "string") {
      state.noteText = saved.noteText;
      elements.notesPad.value = saved.noteText;
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
    elements.plannerMeta.textContent = `플래너: ${plan.meta?.planner || "-"}${plan.meta?.reason ? ` | 사유: ${plan.meta.reason}` : ""}`;

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
  restoreWorkspace();
  if (state.sheetRows.length === 0) {
    resetSheetData();
  }
  renderSheet();
  renderSlides();
  await refreshDocumentView();
  setMode(state.mode || "writer");
  setStatus(
    "준비 완료",
    "HWP/HWPX 파일을 열거나 오른쪽 에이전트에 업무 문서 요청을 입력하면 됩니다.",
  );
  setBadge("준비 완료");
}

elements.tabWriter.addEventListener("click", () => setMode("writer"));
elements.tabNotes.addEventListener("click", () => setMode("notes"));
elements.tabSheet.addEventListener("click", () => setMode("sheet"));
elements.tabSlides.addEventListener("click", () => setMode("slides"));

elements.newDocument.addEventListener("click", async () => {
  createBlankDocument();
  await refreshDocumentView();
  setStatus("빈 문서를 새로 만들었습니다.");
  setBadge("새 문서");
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

elements.generateSlides.addEventListener("click", () => {
  generateSlidesFromPrompt(elements.promptInput.value);
  setMode("slides");
  setStatus("현재 요청을 기준으로 슬라이드 초안을 생성했습니다.");
});

elements.exportSlides.addEventListener("click", () => {
  exportSlidesMarkdown();
  setStatus("슬라이드 초안을 Markdown으로 저장했습니다.");
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
