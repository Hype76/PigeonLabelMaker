const state = {
  settings: null,
  profiles: [],
  builtInFonts: [],
  availableFonts: [],
  useSystemFonts: false,
  imagePath: "",
  serialPorts: [],
  bleDevices: [],
  bleWritable: [],
  bleNotify: [],
  bleState: { connected: false, address: "", name: "" },
  batteryLevel: null,
  comConnected: false,
  comConnecting: false,
  isHydrating: false,
  previewTimer: null,
  saveTimer: null,
  busy: false,
  printQueueActive: false,
  printStopRequested: false,
  historyPast: [],
  historyFuture: [],
  connectionTestStatus: "idle",
  connecting: false,
  connectionStartTime: null,
  connectionTimerInterval: null,
  isCanvasInteracting: false,
  canvasItems: [],
  selectedItem: null,
  dragContext: null,
  canvasBackground: "",
  inlineEditor: null,
  editingItemIndex: null,
  lastCanvasClickIndex: null,
  lastCanvasClickAt: 0,
  savedDesigns: [],
};

const elements = {};
const APP_NAME = "Pigeon Label Maker";

document.title = APP_NAME;

function $(id) {
  return document.getElementById(id);
}

function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem("theme", theme);
}

function loadTheme() {
  const saved = localStorage.getItem("theme") || "dark";
  applyTheme(saved);
}

function toggleTheme() {
  const current = document.documentElement.getAttribute("data-theme") || "dark";
  applyTheme(current === "dark" ? "light" : "dark");
}

function applyMode(mode) {
  document.documentElement.setAttribute("data-mode", mode);
  localStorage.setItem("ui_mode", mode);

  if (elements.modeToggle) {
    elements.modeToggle.textContent = mode === "advanced" ? "Simple" : "Advanced";
  }
}

function loadMode() {
  const saved = localStorage.getItem("ui_mode") || "simple";
  applyMode(saved);
}

function toggleMode() {
  const current = document.documentElement.getAttribute("data-mode") || "simple";
  applyMode(current === "simple" ? "advanced" : "simple");
}

function checkOnboarding() {
  const seen = localStorage.getItem("onboarding_seen");

  if (!seen) {
    elements.onboardingOverlay.classList.remove("hidden");
  }
}

function closeOnboarding() {
  elements.onboardingOverlay.classList.add("hidden");
  localStorage.setItem("onboarding_seen", "true");
}

function openHelp() {
  elements.helpOverlay.classList.remove("hidden");
}

function closeHelp() {
  elements.helpOverlay.classList.add("hidden");
}

function applySidebarSectionState(section, expanded) {
  section.classList.toggle("collapsed", !expanded);
  const toggle = section.querySelector(".section-toggle");
  if (toggle) {
    toggle.setAttribute("aria-expanded", expanded ? "true" : "false");
  }
}

function loadSidebarSectionState(sectionId) {
  const saved = localStorage.getItem(`sidebar_section_${sectionId}`);
  if (saved) {
    return saved !== "collapsed";
  }
  return sectionId === "brand" || sectionId === "designer";
}

function initSidebarCollapsibles() {
  document.querySelectorAll(".sidebar-section").forEach((section) => {
    const sectionId = section.dataset.sectionId;
    const toggle = section.querySelector(".section-toggle");
    if (!sectionId || !toggle) {
      return;
    }

    applySidebarSectionState(section, loadSidebarSectionState(sectionId));
    toggle.addEventListener("click", () => {
      const expanded = section.classList.contains("collapsed");
      applySidebarSectionState(section, expanded);
      localStorage.setItem(
        `sidebar_section_${sectionId}`,
        expanded ? "expanded" : "collapsed"
      );
    });
  });
}

async function loadFonts() {
  const fontSet = new Set(state.builtInFonts || []);

  if (state.useSystemFonts) {
    if (typeof window.queryLocalFonts === "function") {
      try {
        const localFonts = await window.queryLocalFonts();
        for (const font of localFonts) {
          if (font?.family) {
            fontSet.add(font.family);
          }
        }
      } catch (_error) {
      }
    }

    if (document.fonts?.forEach) {
      document.fonts.forEach((font) => {
        if (font?.family) {
          fontSet.add(font.family.replaceAll("\"", ""));
        }
      });
    }
  }

  state.availableFonts = [...fontSet].sort((left, right) => left.localeCompare(right));
  if (state.availableFonts.length === 0) {
    state.availableFonts = ["Arial Bold"];
  }

  fillSelect(elements.fontPicker, state.availableFonts, (item) => item, (item) => item);
  syncSelectedItemControls();
}

function humanizeErrorMessage(message) {
  const text = String(message || "").trim();
  const lower = text.toLowerCase();

  if (!text) {
    return "Something went wrong. Try again.";
  }

  if (lower.includes("could not open port")) {
    const match = text.match(/port '([^']+)'/i);
    const port = match?.[1] || state.settings?.port || "the selected port";
    return `${port} could not be opened. Check that the printer is connected and that the port still exists.`;
  }

  if (lower.includes("device which does not exist")) {
    return "The selected printer is no longer available. Refresh the ports list and choose the correct device again.";
  }

  if (lower.includes("select a com port first")) {
    return "Choose a COM port, then click Connect COM.";
  }

  if (lower.includes("connect com port first")) {
    return "Click Connect COM before printing.";
  }

  if (lower.includes("select a ble device first")) {
    return "Choose a BLE printer, then click Connect.";
  }

  if (lower.includes("connect to ble device first")) {
    return "Connect to the BLE printer before printing.";
  }

  if (lower.includes("backend crashed") || lower.includes("backend exited")) {
    return "The print service restarted. Try the action again.";
  }

  if (lower.startsWith("error invoking remote method")) {
    const cleaned = text.replace(/^Error invoking remote method '[^']+':\s*/i, "");
    return humanizeErrorMessage(cleaned);
  }

  return text;
}

function showError(message) {
  elements.errorBar.textContent = humanizeErrorMessage(message);
  elements.errorBar.classList.remove("hidden");
}

function clearError() {
  elements.errorBar.classList.add("hidden");
  elements.errorBar.textContent = "";
}

function validateBeforePrint() {
  const hasContent = state.canvasItems.some((item) => {
    if (item.type === "image") {
      return Boolean(item.imageData);
    }
    return Boolean(String(item.text || "").trim());
  });

  if (!hasContent) {
    return "Add at least one design element";
  }

  if (state.settings.output_mode === "Printer") {
    if (!state.settings.port) {
      return "Select a COM port";
    }
    if (!state.comConnected) {
      return "Connect COM port first";
    }
  }

  if (state.settings.output_mode === "BLE") {
    if (!state.settings.ble_device_address) {
      return "Select a BLE device";
    }

    if (!state.bleState.connected) {
      return "Connect to BLE device first";
    }
  }

  return null;
}

function setupShortcuts() {
  document.addEventListener("keydown", (e) => {
    const targetTag = String(e.target?.tagName || "").toUpperCase();
    const isTypingTarget =
      targetTag === "INPUT" ||
      targetTag === "TEXTAREA" ||
      targetTag === "SELECT" ||
      Boolean(e.target?.isContentEditable);

    if (e.ctrlKey && e.key === "Enter") {
      e.preventDefault();
      runPrint();
    }

    if (e.ctrlKey && e.key.toLowerCase() === "l") {
      e.preventDefault();
      elements.layer1Text.focus();
    }

    if (e.ctrlKey && e.key === "Backspace") {
      e.preventDefault();
      const item = state.selectedItem !== null ? state.canvasItems[state.selectedItem] : state.canvasItems[0];
      if (item && item.type !== "image") {
        item.text = "";
        syncCanvasIntoSettings();
        renderCanvasFromState();
      }
      queuePreview();
      queueSave();
    }

    if (e.ctrlKey && e.key.toLowerCase() === "t") {
      e.preventDefault();
      toggleTheme();
    }

    if (e.ctrlKey && e.key.toLowerCase() === "m") {
      e.preventDefault();
      toggleMode();
    }

    if (e.ctrlKey && e.key.toLowerCase() === "z" && !e.shiftKey) {
      e.preventDefault();
      undoCanvas();
    }

    if ((e.ctrlKey && e.key.toLowerCase() === "y") || (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === "z")) {
      e.preventDefault();
      redoCanvas();
    }

    if (e.ctrlKey && e.key.toLowerCase() === "d" && !isTypingTarget) {
      e.preventDefault();
      duplicateSelectedItem();
    }

    if (!isTypingTarget && e.key === "Delete" && state.selectedItem !== null) {
      e.preventDefault();
      deleteItem(state.selectedItem);
    }

    if (!isTypingTarget && state.selectedItem !== null && e.key.startsWith("Arrow")) {
      e.preventDefault();
      const item = state.canvasItems[state.selectedItem];
      if (!item || item.locked) {
        return;
      }
      const step = e.shiftKey ? 10 : 1;
      if (e.key === "ArrowLeft") {
        item.x -= step;
      } else if (e.key === "ArrowRight") {
        item.x += step;
      } else if (e.key === "ArrowUp") {
        item.y -= step;
      } else if (e.key === "ArrowDown") {
        item.y += step;
      }
      clampItemToCanvas(item, elements.designCanvas);
      syncCanvasIntoSettings();
      renderCanvasFromState();
      recordCanvasHistory();
      queueSave();
      triggerPreviewAfterInteraction();
    }
  });
}

function saveRecentLabel(text) {
  if (!text.trim()) {
    return;
  }

  let items = JSON.parse(localStorage.getItem("recent_labels") || "[]");

  items = items.filter((item) => item !== text);
  items.unshift(text);

  if (items.length > 10) {
    items = items.slice(0, 10);
  }

  localStorage.setItem("recent_labels", JSON.stringify(items));
  renderRecentLabels();
}

function nextCanvasId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function mmToPixels(mm, dpi) {
  return Math.max(1, Math.round((Number(mm) / 25.4) * Number(dpi)));
}

function getCanvasLogicalSize() {
  const dpi = Number(state.settings?.print_dpi) || 203;
  return {
    width: mmToPixels(Number(state.settings?.label_width_mm) || 40, dpi),
    height: mmToPixels(Number(state.settings?.label_height_mm) || 14, dpi),
  };
}

function getCanvasPreviewScale() {
  return Number(elements.designCanvas?.dataset.previewScale) || 1;
}

function makeCanvasItem(type, overrides = {}) {
  const defaults = {
    id: nextCanvasId(),
    key: "",
    type,
    text:
      type === "text"
        ? "New Text"
        : type === "qr"
          ? "QR DATA"
          : type === "barcode"
            ? "123456789"
            : "",
    x: 50,
    y: 50,
    width: type === "qr" || type === "image" ? 120 : 150,
    height: type === "qr" || type === "image" ? 120 : 60,
    font: state.settings?.font_name || "Arial Bold",
    fontSize: type === "text" ? 50 : null,
    fittedFontSize: null,
    fontWeight: type === "text" ? 700 : 400,
    textTransform: "none",
    rotation: 0,
    invert: false,
    imageData: "",
    aspectRatio: type === "image" ? 1 : null,
    locked: false,
  };
  const item = { ...defaults, ...overrides };
  if (item.type === "image") {
    const width = Math.max(1, Number(item.width) || defaults.width);
    const height = Math.max(1, Number(item.height) || defaults.height);
    item.aspectRatio = Number(item.aspectRatio) || width / height;
  }
  return item;
}

function formatExpiryDate(value) {
  const parts = String(value || "").split("-");
  if (parts.length !== 3) {
    return "";
  }

  const [year, month, day] = parts;
  if (!year || !month || !day) {
    return "";
  }

  return `${day}/${month}/${year}`;
}

function buildExpiryCanvasItems(dateText) {
  const canvasWidth = elements.designCanvas?.clientWidth || 600;
  const canvasHeight = elements.designCanvas?.clientHeight || 210;
  const sidePadding = Math.max(20, Math.round(canvasWidth * 0.08));
  const availableWidth = Math.max(120, canvasWidth - (sidePadding * 2));
  const titleHeight = Math.max(28, Math.round(canvasHeight * 0.2));
  const dateHeight = Math.max(34, Math.round(canvasHeight * 0.26));
  const titleY = Math.max(18, Math.round(canvasHeight * 0.18));
  const gap = Math.max(8, Math.round(canvasHeight * 0.08));
  const dateY = titleY + titleHeight + gap;

  return [
    makeCanvasItem("text", {
      key: "layer1",
      text: "Expires",
      x: sidePadding,
      y: titleY,
      width: availableWidth,
      height: titleHeight,
      fontSize: Math.max(18, Math.round(canvasHeight * 0.16)),
    }),
    makeCanvasItem("text", {
      key: "layer2",
      text: dateText,
      x: sidePadding,
      y: dateY,
      width: availableWidth,
      height: dateHeight,
      fontSize: Math.max(22, Math.round(canvasHeight * 0.22)),
    }),
  ];
}

function renderRecentLabels() {
  const container = elements.recentLabels;
  container.innerHTML = "";

  const items = JSON.parse(localStorage.getItem("recent_labels") || "[]");

  items.forEach((text) => {
    const chip = document.createElement("div");
    chip.className = "recent-chip";
    chip.textContent = text;

    chip.addEventListener("click", () => {
      const target = state.selectedItem !== null
        ? state.canvasItems[state.selectedItem]
        : state.canvasItems.find((item) => item.type === "text");
      if (!target) {
        return;
      }
      target.text = text;
      syncCanvasIntoSettings();
      syncSelectedItemControls();
      renderCanvasFromState();
      queuePreview();
      queueSave();
    });

    container.appendChild(chip);
  });
}

function addPrintHistory(success, message) {
  const time = new Date().toLocaleTimeString("en-GB");

  let items = JSON.parse(localStorage.getItem("print_history") || "[]");

  items.unshift({
    time,
    success,
    message,
  });

  if (items.length > 20) {
    items = items.slice(0, 20);
  }

  localStorage.setItem("print_history", JSON.stringify(items));
  renderPrintHistory();
}

function renderPrintHistory() {
  const container = elements.printHistory;
  container.innerHTML = "";

  const items = JSON.parse(localStorage.getItem("print_history") || "[]");

  items.forEach((item) => {
    const div = document.createElement("div");
    div.className = `history-item ${item.success ? "history-success" : "history-fail"}`;
    div.textContent = `[${item.time}] ${item.message}`;
    container.appendChild(div);
  });
}

function cloneCanvasItems(items = state.canvasItems) {
  return structuredClone(
    items.map((item) => ({
      ...item,
      _tempX: undefined,
      _tempY: undefined,
      _tempWidth: undefined,
      _tempHeight: undefined,
    }))
  );
}

function snapshotCanvasState() {
  return JSON.stringify(cloneCanvasItems());
}

function resetCanvasHistory() {
  state.historyPast = [snapshotCanvasState()];
  state.historyFuture = [];
}

function recordCanvasHistory() {
  const snapshot = snapshotCanvasState();
  if (state.historyPast[state.historyPast.length - 1] === snapshot) {
    return;
  }
  state.historyPast.push(snapshot);
  if (state.historyPast.length > 80) {
    state.historyPast = state.historyPast.slice(-80);
  }
  state.historyFuture = [];
  updateCanvasActionButtons();
}

function restoreCanvasSnapshot(snapshot) {
  state.canvasItems = JSON.parse(snapshot).map((item) =>
    makeCanvasItem(String(item.type || "text").toLowerCase(), item)
  );
  if (state.selectedItem !== null && state.selectedItem >= state.canvasItems.length) {
    state.selectedItem = state.canvasItems.length ? state.canvasItems.length - 1 : null;
  }
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  queuePreview();
  updateCanvasActionButtons();
}

function undoCanvas() {
  if (state.historyPast.length <= 1) {
    return;
  }
  const current = state.historyPast.pop();
  state.historyFuture.push(current);
  restoreCanvasSnapshot(state.historyPast[state.historyPast.length - 1]);
  setStatus("Undo");
}

function redoCanvas() {
  if (state.historyFuture.length === 0) {
    return;
  }
  const snapshot = state.historyFuture.pop();
  state.historyPast.push(snapshot);
  restoreCanvasSnapshot(snapshot);
  setStatus("Redo");
}

function updateCanvasActionButtons() {
  const selected = state.selectedItem !== null ? state.canvasItems[state.selectedItem] : null;
  const isText = selected?.type === "text";
  if (elements.undoButton) {
    elements.undoButton.disabled = state.historyPast.length <= 1;
  }
  if (elements.redoButton) {
    elements.redoButton.disabled = state.historyFuture.length === 0;
  }
  if (elements.duplicateButton) {
    elements.duplicateButton.disabled = !selected;
  }
  if (elements.lockButton) {
    elements.lockButton.disabled = !selected;
    elements.lockButton.textContent = selected?.locked ? "Unlock" : "Lock";
    elements.lockButton.classList.toggle("is-active", Boolean(selected?.locked));
  }
  if (elements.sendBackwardButton) {
    elements.sendBackwardButton.disabled = !selected || state.selectedItem === 0;
  }
  if (elements.bringForwardButton) {
    elements.bringForwardButton.disabled = !selected || state.selectedItem === state.canvasItems.length - 1;
  }
  if (elements.fitTextButton) {
    elements.fitTextButton.disabled = !isText;
  }
  if (elements.toggleBoldButton) {
    elements.toggleBoldButton.disabled = !isText;
    elements.toggleBoldButton.classList.toggle("is-active", Boolean(selected && Number(selected.fontWeight) >= 700));
  }
  if (elements.toggleUppercaseButton) {
    elements.toggleUppercaseButton.disabled = !isText;
    elements.toggleUppercaseButton.classList.toggle("is-active", Boolean(selected && selected.textTransform === "uppercase"));
  }
}

function loadSavedDesigns() {
  try {
    state.savedDesigns = JSON.parse(localStorage.getItem("saved_designs") || "[]");
  } catch (_error) {
    state.savedDesigns = [];
  }
}

function renderSavedDesigns() {
  if (!elements.savedDesigns) {
    return;
  }
  elements.savedDesigns.innerHTML = "";
  for (const design of state.savedDesigns) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "saved-design-card ghost";
    button.innerHTML = `
      <img src="${design.thumbnail}" alt="${design.name}" />
      <span>${design.name}</span>
    `;
    button.addEventListener("click", () => {
      state.canvasItems = design.items.map((item) =>
        makeCanvasItem(String(item.type || "text").toLowerCase(), item)
      );
      state.selectedItem = state.canvasItems.length ? 0 : null;
      syncCanvasIntoSettings();
      renderCanvasFromState();
      resetCanvasHistory();
      queueSave();
      queuePreview();
      setStatus(`Loaded ${design.name}`);
    });
    elements.savedDesigns.appendChild(button);
  }
}

function updateConnectionTestBadge(status = state.connectionTestStatus) {
  state.connectionTestStatus = status;
  if (!elements.connectionTestBadge) {
    return;
  }
  const label =
    status === "pass" ? "Connection OK" :
    status === "fail" ? "Connection Failed" :
    "Not tested";
  elements.connectionTestBadge.textContent = label;
  elements.connectionTestBadge.className = `status-pill${status === "idle" ? " muted" : ""}`;
}

function showPrintSuccess() {
  elements.printButton.classList.add("print-success");

  setTimeout(() => {
    elements.printButton.classList.remove("print-success");
  }, 400);
}

function canvasItemText(item) {
  if (item.type === "qr") {
    return item.text || "[QR]";
  }
  if (item.type === "barcode") {
    return item.text || "[Barcode]";
  }
  return item.text || "Text";
}

function normalizeInlineEditorText(value) {
  return String(value || "")
    .replaceAll("\r", "")
    .replaceAll("\n", "");
}

function insertTextAtSelection(text) {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return false;
  }

  const range = selection.getRangeAt(0);
  range.deleteContents();
  const node = document.createTextNode(text);
  range.insertNode(node);
  range.setStartAfter(node);
  range.setEndAfter(node);
  selection.removeAllRanges();
  selection.addRange(range);
  return true;
}

function escapeSvgText(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;");
}

function fallbackQrSvgMarkup(text) {
  const content = text || "QR";
  let hash = 0;
  for (const char of content) {
    hash = ((hash << 5) - hash + char.charCodeAt(0)) >>> 0;
  }

  const cells = [];
  const size = 17;
  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      const bit = (hash >> ((x + (y * size)) % 24)) & 1;
      const finder =
        ((x < 5 && y < 5) || (x > 11 && y < 5) || (x < 5 && y > 11))
        && !(x === 4 || y === 4);
      if (bit || finder) {
        cells.push(`<rect x="${x}" y="${y}" width="1" height="1" fill="#111111" />`);
      }
    }
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${size} ${size}" shape-rendering="crispEdges"><rect width="${size}" height="${size}" fill="#ffffff"/>${cells.join("")}</svg>`;
}

function qrSvgMarkup(text) {
  const content = String(text || "QR");
  try {
    const svg = window.pigeonApi?.renderQrSvg?.(content);
    if (svg) {
      return svg;
    }
  } catch (_error) {
  }
  return fallbackQrSvgMarkup(content);
}

function qrSvgDataUri(text) {
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(qrSvgMarkup(text))}`;
}

function barcodeSvgMarkup(text) {
  const content = String(text || "123456789");
  try {
    const svg = window.pigeonApi?.renderBarcodeSvg?.(content);
    if (svg) {
      return svg;
    }
  } catch (_error) {
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 180 68"><rect width="100%" height="100%" fill="#ffffff"/><text x="50%" y="52%" text-anchor="middle" font-family="Arial" font-size="12" fill="#111111">${escapeSvgText(content)}</text></svg>`;
}

function barcodeSvgDataUri(text) {
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(barcodeSvgMarkup(text))}`;
}

function generatePreviewImage(item) {
  if (item.type === "qr") {
    return qrSvgDataUri(item.text);
  }
  if (item.type === "barcode") {
    return barcodeSvgDataUri(item.text);
  }
  if (item.type === "image") {
    return item.imageData || "";
  }
  return "";
}

function generatePreviewMarkup(item) {
  if (item.type === "qr") {
    return qrSvgMarkup(item.text);
  }
  if (item.type === "barcode") {
    return barcodeSvgMarkup(item.text);
  }
  return "";
}

function syncSelectedItemControls() {
  const item = state.selectedItem !== null ? state.canvasItems[state.selectedItem] : state.canvasItems[0];
  if (!item) {
    elements.layer1Text.value = "";
    elements.fontSizeInput.value = "";
    return;
  }

  if (state.availableFonts.length > 0) {
    const selectedFont = item.font || state.settings.font_name || state.availableFonts[0];
    if (!state.availableFonts.includes(selectedFont)) {
      state.availableFonts = [selectedFont, ...state.availableFonts].sort((left, right) => left.localeCompare(right));
      fillSelect(elements.fontPicker, state.availableFonts, (font) => font, (font) => font);
    }
    elements.fontPicker.value = selectedFont;
  }

  elements.fontSizeInput.value = String(Math.max(6, Number(item.fontSize) || 50));
  elements.layer1Text.value = item.text || "";
  elements.layer1Text.disabled = item.type === "image";
  elements.fontPicker.disabled = item.type !== "text";
  elements.fontSizeInput.disabled = item.type !== "text";
  elements.layer1Mode.value =
    item.type === "text" ? "Text" : item.type === "qr" ? "QR" : item.type === "barcode" ? "Barcode" : "Text";
  updateCanvasActionButtons();
}

function commitSelectedFontSize(rawValue, options = {}) {
  const { restoreIfEmpty = false } = options;
  if (state.selectedItem === null) {
    return;
  }

  const item = state.canvasItems[state.selectedItem];
  if (!item || item.type !== "text") {
    return;
  }

  const value = String(rawValue ?? "").trim();
  if (!value) {
    if (restoreIfEmpty) {
      const fallback = Math.max(6, Number(item.fontSize) || 50);
      elements.fontSizeInput.value = String(fallback);
    }
    return;
  }

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    if (restoreIfEmpty) {
      elements.fontSizeInput.value = String(Math.max(6, Number(item.fontSize) || 50));
    }
    return;
  }

  const size = Math.max(6, numeric);
  item.fontSize = size;
  elements.fontSizeInput.value = String(size);
  clearError();
  syncCanvasIntoSettings();
  renderCanvasFromState();
  queueSave();
  queuePreview();
}

function syncCanvasIntoSettings() {
  const canvasWidth = elements.designCanvas.clientWidth || 1;
  const canvasHeight = elements.designCanvas.clientHeight || 1;

  state.settings.canvas_layout = state.canvasItems.map((item) => ({
    id: item.id,
    key: item.key || "",
    type: item.type,
    text: item.text || "",
    x: item.x / canvasWidth,
    y: item.y / canvasHeight,
    width: item.width / canvasWidth,
    height: item.height / canvasHeight,
    font: item.font || state.settings.font_name,
    fontScale: item.fontSize ? (item.fontSize / canvasHeight) : null,
    fontWeight: Number(item.fontWeight) || 400,
    textTransform: item.textTransform || "none",
    rotation: Number(item.rotation) || 0,
    locked: Boolean(item.locked),
    invert: Boolean(item.invert),
    imageData: item.imageData || "",
  }));

  const printableItems = state.canvasItems.filter((item) => item.type !== "image");
  const layer1Item = printableItems[0];
  const layer2Item = printableItems[1];

  state.settings.layer1.text = layer1Item?.text || "";
  state.settings.layer1.mode = layer1Item
    ? layer1Item.type === "qr"
      ? "QR"
      : layer1Item.type === "barcode"
        ? "Barcode"
        : "Text"
    : "Text";

  state.settings.layer2.text = layer2Item?.text || "";
  state.settings.layer2.mode = layer2Item
    ? layer2Item.type === "qr"
      ? "QR"
      : layer2Item.type === "barcode"
        ? "Barcode"
        : "Text"
    : "Off";
}

function buildRasterRenderSettings(overrides = {}) {
  return {
    ...state.settings,
    ...overrides,
    canvas_layout: [],
    layer1: {
      ...state.settings.layer1,
      text: "",
      mode: "Off",
    },
    layer2: {
      ...state.settings.layer2,
      text: "",
      mode: "Off",
    },
  };
}

function normalizeCanvasFontFamily(fontName) {
  const trimmed = String(fontName || "").trim();
  if (!trimmed) {
    return "\"Arial\", sans-serif";
  }

  const escaped = trimmed.replaceAll("\"", "\\\"");
  return `"${escaped}", sans-serif`;
}

async function ensureCanvasFontsReady() {
  if (!document.fonts) {
    return;
  }

  const fontFamilies = [...new Set(
    state.canvasItems
      .filter((item) => item.type === "text" && String(item.text || "").trim())
      .map((item) => normalizeCanvasFontFamily(item.font || state.settings?.font_name || "Arial"))
  )];

  if (fontFamilies.length === 0) {
    return;
  }

  await Promise.all(fontFamilies.map(async (fontFamily) => {
    try {
      await document.fonts.load(`32px ${fontFamily}`);
      await document.fonts.load(`64px ${fontFamily}`);
    } catch (_error) {
    }
  }));

  try {
    await document.fonts.ready;
  } catch (_error) {
  }
}

async function captureCanvasImage() {
  const canvasEl = elements.designCanvas;
  if (typeof window.html2canvas !== "function") {
    throw new Error("Canvas capture is not available");
  }

  await ensureCanvasFontsReady();
  const captureHost = document.createElement("div");
  captureHost.style.position = "fixed";
  captureHost.style.left = "-100000px";
  captureHost.style.top = "0";
  captureHost.style.pointerEvents = "none";
  captureHost.style.opacity = "0";
  captureHost.style.zIndex = "-1";

  const clone = canvasEl.cloneNode(true);
  clone.classList.add("capture-render");
  clone.style.transform = "none";
  clone.style.transition = "none";
  clone.dataset.previewScale = "1";
  captureHost.appendChild(clone);
  document.body.appendChild(captureHost);

  try {
    await new Promise((resolve) => window.requestAnimationFrame(resolve));
    await new Promise((resolve) => window.requestAnimationFrame(resolve));

    const canvas = await window.html2canvas(clone, {
      backgroundColor: "#ffffff",
      scale: 1,
      useCORS: true,
      ignoreElements: (element) =>
        Boolean(
          element?.classList?.contains("canvas-safe") ||
          element?.classList?.contains("guide-line") ||
          element?.classList?.contains("resize-handle") ||
          element?.classList?.contains("rotate-handle") ||
          element?.classList?.contains("rotate-line") ||
          element?.classList?.contains("edit-btn") ||
          element?.classList?.contains("delete-btn")
        ),
    });

    return canvas.toDataURL("image/png");
  } finally {
    captureHost.remove();
  }
}

function fitTextToBox(el, item) {
  let fontSize = item.height * 0.6;
  el.style.fontSize = `${fontSize}px`;

  while ((el.scrollWidth > item.width || el.scrollHeight > item.height) && fontSize > 5) {
    fontSize -= 1;
    el.style.fontSize = `${fontSize}px`;
  }

  item.fittedFontSize = fontSize;
}

function buildCanvasItemsFromSettings() {
  if (state.canvasItems.length > 0) {
    return;
  }

  const canvasWidth = elements.designCanvas?.clientWidth || 600;
  const canvasHeight = elements.designCanvas?.clientHeight || 320;
  if (Array.isArray(state.settings.canvas_layout) && state.settings.canvas_layout.length > 0) {
    state.canvasItems = state.settings.canvas_layout.map((item, index) => makeCanvasItem(
      String(item.type || "text").toLowerCase(),
      {
        id: item.id || nextCanvasId(),
        key: item.key || "",
        text: item.text || "",
        x: Number(item.x) <= 1.5 ? Math.round(Number(item.x || 0) * canvasWidth) : Number(item.x || 0),
        y: Number(item.y) <= 1.5 ? Math.round(Number(item.y || 0) * canvasHeight) : Number(item.y || 0),
        width: Number(item.width) <= 1.5 ? Math.round(Number(item.width || 0.25) * canvasWidth) : Number(item.width || 150),
        height: Number(item.height) <= 1.5 ? Math.round(Number(item.height || 0.2) * canvasHeight) : Number(item.height || 60),
        font: item.font || state.settings.font_name,
        fontSize: item.fontScale ? Math.round(Number(item.fontScale) * canvasHeight) : null,
        fontWeight: Number(item.fontWeight) || 400,
        textTransform: item.textTransform || "none",
        rotation: Number(item.rotation) || 0,
        locked: Boolean(item.locked),
        invert: Boolean(item.invert),
        imageData: item.imageData || "",
      }
    ));
  } else {
    state.canvasItems = [
      makeCanvasItem("text", {
        key: "layer1",
        text: state.settings.layer1.text || "Label",
        x: 40,
        y: 40,
        width: 200,
        height: 60,
      }),
    ];
    if (state.settings.layer2.mode && state.settings.layer2.mode !== "Off") {
      state.canvasItems.push(makeCanvasItem(
        String(state.settings.layer2.mode).toLowerCase(),
        {
          key: "layer2",
          text: state.settings.layer2.text || "",
          x: 260,
          y: 40,
          width: state.settings.layer2.mode === "Text" ? 180 : 120,
          height: state.settings.layer2.mode === "Text" ? 60 : 120,
        }
      ));
    }
  }

  if (state.selectedItem !== null && state.selectedItem >= state.canvasItems.length) {
    state.selectedItem = null;
  }
}

function updateCanvasSurface() {
  const logical = getCanvasLogicalSize();
  const viewportWidth = elements.designCanvasViewport?.clientWidth || logical.width;
  const viewportHeight = elements.designCanvasViewport?.clientHeight || logical.height;
  const scale = Math.max(
    0.1,
    Math.min(viewportWidth / logical.width, viewportHeight / logical.height)
  );

  elements.designCanvas.style.width = `${logical.width}px`;
  elements.designCanvas.style.height = `${logical.height}px`;
  elements.designCanvas.style.transform = `scale(${scale})`;
  elements.designCanvas.dataset.previewScale = String(scale);
}

function updatePrintPreviewOrientation() {
  elements.printPreview.style.rotate = "0deg";
}

function getRenderedItemWidth(item, el = null) {
  return el?.offsetWidth || Math.max(40, Number(item.width) || 0);
}

function getRenderedItemHeight(item, el = null) {
  return el?.offsetHeight || Math.max(20, Number(item.height) || 0);
}

function getCanvasBounds(canvas, item, el = null) {
  const canvasWidth = canvas.clientWidth || 0;
  const canvasHeight = canvas.clientHeight || 0;
  const itemWidth = getRenderedItemWidth(item, el);
  const itemHeight = getRenderedItemHeight(item, el);
  const safeMargin = 10;
  const visibleGrab = 14;

  return {
    canvasWidth,
    canvasHeight,
    itemWidth,
    itemHeight,
    safeMargin,
    visibleGrab,
    safeMinX: safeMargin,
    safeMinY: safeMargin,
    safeMaxX: canvasWidth - itemWidth - safeMargin,
    safeMaxY: canvasHeight - itemHeight - safeMargin,
    hardMinX: visibleGrab - itemWidth,
    hardMinY: visibleGrab - itemHeight,
    hardMaxX: canvasWidth - visibleGrab,
    hardMaxY: canvasHeight - visibleGrab,
  };
}

function clampItemToCanvas(item, canvas, el = null) {
  const bounds = getCanvasBounds(canvas, item, el);

  if (item.x < bounds.hardMinX) {
    item.x = bounds.hardMinX;
  }
  if (item.y < bounds.hardMinY) {
    item.y = bounds.hardMinY;
  }

  if (item.x > bounds.hardMaxX) {
    item.x = bounds.hardMaxX;
  }

  if (item.y > bounds.hardMaxY) {
    item.y = bounds.hardMaxY;
  }
}

function applyStickyBounds(item, canvas, el = null) {
  const bounds = getCanvasBounds(canvas, item, el);
  let hitEdge = false;

  if (item.x < bounds.safeMinX) {
    item.x = bounds.safeMinX + (item.x - bounds.safeMinX) * 0.3;
    hitEdge = true;
  }

  if (item.x > bounds.safeMaxX) {
    item.x = bounds.safeMaxX + (item.x - bounds.safeMaxX) * 0.3;
    hitEdge = true;
  }

  if (item.y < bounds.safeMinY) {
    item.y = bounds.safeMinY + (item.y - bounds.safeMinY) * 0.3;
    hitEdge = true;
  }

  if (item.y > bounds.safeMaxY) {
    item.y = bounds.safeMaxY + (item.y - bounds.safeMaxY) * 0.3;
    hitEdge = true;
  }

  if (item.x < bounds.hardMinX) {
    item.x = bounds.hardMinX;
  }
  if (item.x > bounds.hardMaxX) {
    item.x = bounds.hardMaxX;
  }
  if (item.y < bounds.hardMinY) {
    item.y = bounds.hardMinY;
  }
  if (item.y > bounds.hardMaxY) {
    item.y = bounds.hardMaxY;
  }

  if (hitEdge) {
    elements.canvasSafe.classList.add("canvas-edge-hit");
  } else {
    elements.canvasSafe.classList.remove("canvas-edge-hit");
  }
}

function renderCanvasFromState() {
  const canvas = elements.designCanvas;
  state.inlineEditor = null;
  canvas.querySelectorAll(".canvas-item").forEach((item) => item.remove());
  canvas.style.backgroundImage = "none";

  state.canvasItems.forEach((item, index) => {
    const el = document.createElement("div");
    el.className = "canvas-item";
    el.dataset.type = item.type;
    el.style.left = `${item.x}px`;
    el.style.top = `${item.y}px`;
    el.style.width = `${getRenderedItemWidth(item)}px`;
    el.style.height = `${getRenderedItemHeight(item)}px`;
    el.style.display = "flex";
    el.style.alignItems = "center";
    el.style.justifyContent = "center";
    el.style.whiteSpace = item.type === "text" ? "pre-wrap" : "nowrap";
    el.style.overflow = "hidden";
    el.style.textAlign = "center";
    el.style.lineHeight = "1.05";
    el.style.padding = item.type === "text" ? "4px" : "0";
    el.style.pointerEvents = "auto";
    el.style.transform = `rotate(${item.rotation || 0}deg)`;
    el.style.transformOrigin = "center center";
    el.classList.toggle("locked", Boolean(item.locked));

    if (item.type === "text") {
      el.innerText = canvasItemText(item);
      el.style.fontFamily = normalizeCanvasFontFamily(item.font || state.settings.font_name);
      el.style.textRendering = "geometricPrecision";
      el.style.fontWeight = String(Number(item.fontWeight) || 400);
      el.style.textTransform = item.textTransform || "none";
      const fontSize = Math.max(6, Number(item.fontSize) || 50);
      if (Number.isFinite(fontSize)) {
        el.style.fontSize = `${fontSize}px`;
      } else {
        fitTextToBox(el, item);
      }
    } else if (item.type === "image") {
      const image = document.createElement("img");
      image.src = generatePreviewImage(item);
      image.alt = item.type;
      image.draggable = false;
      image.style.width = "100%";
      image.style.height = "100%";
      image.style.objectFit = "contain";
      el.appendChild(image);
    } else {
      const wrapper = document.createElement("div");
      wrapper.className = "canvas-vector-preview";
      wrapper.innerHTML = generatePreviewMarkup(item);
      el.appendChild(wrapper);
    }

    if (state.editingItemIndex === index && item.type !== "image") {
      el.classList.add("editing");
      el.contentEditable = "true";
      el.spellcheck = false;
      el.innerText = item.text || "";
      el.addEventListener("mousedown", (e) => {
        e.stopPropagation();
      });
      el.addEventListener("click", (e) => {
        e.stopPropagation();
      });
      el.addEventListener("input", () => {
        const nextText = normalizeInlineEditorText(el.innerText);
        if (el.innerText !== nextText) {
          el.innerText = nextText;
          const selection = window.getSelection();
          if (selection) {
            const range = document.createRange();
            range.selectNodeContents(el);
            range.collapse(false);
            selection.removeAllRanges();
            selection.addRange(range);
          }
        }
        item.text = nextText;
        if (state.selectedItem === index) {
          elements.layer1Text.value = nextText;
        }
        syncCanvasIntoSettings();
        queueSave();
        queuePreview();
      });
      el.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
          e.preventDefault();
          closeInlineEditor({ cancel: true });
          return;
        }
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          closeInlineEditor();
        }
      });
      el.addEventListener("blur", () => {
        closeInlineEditor();
      });
      state.inlineEditor = el;
    }

    if (index === state.selectedItem) {
      el.classList.add("selected");
      if (state.editingItemIndex !== index) {
        const handle = document.createElement("button");
        handle.type = "button";
        handle.className = "resize-handle br";
        handle.setAttribute("aria-label", "Resize item");
        handle.title = "Resize";
        handle.innerHTML = `<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M7 13l6-6M10 13h3v-3"/><path d="M13 7H7v6"/></svg>`;
        makeResizable(el, item, handle, index);
        el.appendChild(handle);

        const rotateLine = document.createElement("div");
        rotateLine.className = "rotate-line";

        const rotateHandle = document.createElement("div");
        rotateHandle.className = "rotate-handle";
        rotateHandle.setAttribute("title", "Rotate");
        rotateHandle.setAttribute("aria-label", "Rotate item");
        makeRotatable(el, item, rotateHandle, index);

        el.appendChild(rotateLine);
        el.appendChild(rotateHandle);

        if (item.type !== "image") {
          const edit = document.createElement("button");
          edit.type = "button";
          edit.className = "edit-btn";
          edit.setAttribute("aria-label", "Edit item");
          edit.title = "Edit";
          edit.innerHTML = `<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M4 14.6V16h1.4L14 7.4 12.6 6 4 14.6Z"/><path d="M11.9 6.7 13.3 8.1"/><path d="M12.6 6l1.2-1.2a1.4 1.4 0 0 1 2 0l.4.4a1.4 1.4 0 0 1 0 2L15 8.4"/></svg>`;
          edit.addEventListener("click", (e) => {
            e.stopPropagation();
            e.preventDefault();
            openInlineEditor(index);
          });
          el.appendChild(edit);
        }

        const del = document.createElement("button");
        del.type = "button";
        del.className = "delete-btn";
        del.setAttribute("aria-label", "Delete item");
        del.title = "Delete";
        del.innerHTML = `<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M6 6l8 8"/><path d="M14 6l-8 8"/></svg>`;
        del.addEventListener("click", (e) => {
          e.stopPropagation();
          e.preventDefault();
          deleteItem(index);
        });
        el.appendChild(del);
      }
    }

    el.addEventListener("click", (e) => {
      e.stopPropagation();
      if (state.editingItemIndex === index) {
        return;
      }
      selectItem(index);
    });

      makeDraggable(el, item, index);
      canvas.appendChild(el);
    });

  syncSelectedItemControls();
  if (state.inlineEditor) {
    state.inlineEditor.focus();
    const selection = window.getSelection();
    if (selection) {
      const range = document.createRange();
      range.selectNodeContents(state.inlineEditor);
      range.collapse(false);
      selection.removeAllRanges();
      selection.addRange(range);
    }
  }
}

function deleteItem(index) {
  if (index < 0 || index >= state.canvasItems.length) {
    return;
  }

  state.canvasItems.splice(index, 1);
  state.selectedItem = null;
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  triggerPreviewAfterInteraction();
}

async function saveDesign() {
  localStorage.setItem("design", JSON.stringify(state.canvasItems));
  const thumbnail = await captureCanvasImage();
  const design = {
    id: nextCanvasId(),
    name: `Design ${new Date().toLocaleString("en-GB", {
      day: "2-digit",
      month: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    })}`,
    thumbnail,
    items: cloneCanvasItems(),
  };
  state.savedDesigns.unshift(design);
  state.savedDesigns = state.savedDesigns.slice(0, 12);
  localStorage.setItem("saved_designs", JSON.stringify(state.savedDesigns));
  renderSavedDesigns();
}

function loadDesign() {
  const data = localStorage.getItem("design");
  if (!data) {
    return false;
  }

  try {
    const parsed = JSON.parse(data);
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return false;
    }
    state.canvasItems = parsed.map((item) => makeCanvasItem(String(item.type || "text").toLowerCase(), item));
    state.selectedItem = 0;
    syncCanvasIntoSettings();
    renderCanvasFromState();
    resetCanvasHistory();
    return true;
  } catch (_error) {
    return false;
  }
}

function addItem(type, overrides = {}) {
  const item = makeCanvasItem(type, overrides);
  state.canvasItems.push(item);
  state.selectedItem = state.canvasItems.length - 1;
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  queuePreview();
}

function duplicateSelectedItem() {
  if (state.selectedItem === null) {
    return;
  }
  const source = state.canvasItems[state.selectedItem];
  if (!source) {
    return;
  }
  const duplicate = makeCanvasItem(source.type, {
    ...structuredClone(source),
    id: nextCanvasId(),
    x: Number(source.x) + 12,
    y: Number(source.y) + 12,
  });
  state.canvasItems.splice(state.selectedItem + 1, 0, duplicate);
  state.selectedItem += 1;
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  queuePreview();
}

function toggleSelectedLock() {
  if (state.selectedItem === null) {
    return;
  }
  const item = state.canvasItems[state.selectedItem];
  item.locked = !item.locked;
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  triggerPreviewAfterInteraction();
}

function moveSelectedLayer(direction) {
  if (state.selectedItem === null) {
    return;
  }
  const from = state.selectedItem;
  const to = direction === "forward" ? from + 1 : from - 1;
  if (to < 0 || to >= state.canvasItems.length) {
    return;
  }
  const [item] = state.canvasItems.splice(from, 1);
  state.canvasItems.splice(to, 0, item);
  state.selectedItem = to;
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  triggerPreviewAfterInteraction();
}

function fitSelectedTextToBox() {
  if (state.selectedItem === null) {
    return;
  }
  const item = state.canvasItems[state.selectedItem];
  if (!item || item.type !== "text") {
    return;
  }
  item.fontSize = null;
  renderCanvasFromState();
  const selectedEl = elements.designCanvas.querySelector(".canvas-item.selected");
  if (selectedEl) {
    fitTextToBox(selectedEl, item);
    item.fontSize = Math.max(6, Math.round(item.fittedFontSize || 50));
  }
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  triggerPreviewAfterInteraction();
}

async function exportCurrentDesignPng() {
  try {
    clearError();
    const outputPath = await window.pigeonApi.choosePngPath();
    if (!outputPath) {
      return;
    }
    const imageData = await captureCanvasImage();
    await safeRequest("exportPng", {
      settings: buildRasterRenderSettings(),
      imagePath: imageData,
      outputPath,
    });
    setStatus("PNG exported");
  } catch (error) {
    showError(error.message);
  }
}

function alignSelected(type) {
  if (state.selectedItem === null) {
    return;
  }

  state.isCanvasInteracting = true;
  const item = state.canvasItems[state.selectedItem];
  const canvasWidth = elements.designCanvas.clientWidth || 0;
  const canvasHeight = elements.designCanvas.clientHeight || 0;

  if (type === "left") {
    item.x = 0;
  }
  if (type === "center") {
    item.x = Math.max(0, (canvasWidth - item.width) / 2);
  }
  if (type === "right") {
    item.x = Math.max(0, canvasWidth - item.width);
  }

  if (type === "top") {
    item.y = 0;
  }
  if (type === "middle") {
    item.y = Math.max(0, (canvasHeight - item.height) / 2);
  }
  if (type === "bottom") {
    item.y = Math.max(0, canvasHeight - item.height);
  }

  syncCanvasIntoSettings();
  renderCanvasFromState();
  queueSave();
  setTimeout(() => {
    state.isCanvasInteracting = false;
  }, 50);
  triggerPreviewAfterInteraction();
}

function selectItem(index) {
  state.selectedItem = index;
  const item = state.canvasItems[index];
  if (item?.type === "text") {
    elements.fontSizeInput.value = String(Math.max(6, Number(item.fontSize) || 50));
  }
  renderCanvasFromState();
}

function openInlineEditor(index) {
  const item = state.canvasItems[index];
  if (!item || item.type === "image") {
    return;
  }

  state.selectedItem = index;
  state.lastCanvasClickIndex = null;
  state.lastCanvasClickAt = 0;
  state.editingItemIndex = index;
  renderCanvasFromState();
  requestAnimationFrame(() => {
    state.inlineEditor?.focus();
  });
}

function closeInlineEditor(options = {}) {
  const { cancel = false } = options;
  const index = state.editingItemIndex;
  if (index === null) {
    return;
  }

  const item = state.canvasItems[index];
  const editorValue = normalizeInlineEditorText(state.inlineEditor?.innerText);
  state.editingItemIndex = null;
  state.inlineEditor = null;

  if (!cancel && item) {
    item.text = editorValue ?? item.text;
    syncCanvasIntoSettings();
    syncSelectedItemControls();
    recordCanvasHistory();
    queueSave();
    triggerPreviewAfterInteraction();
  }

  renderCanvasFromState();
}

function insertSymbol(symbol) {
  const item = state.selectedItem !== null ? state.canvasItems[state.selectedItem] : state.canvasItems[0];
  if (!item || item.type === "image") {
    return;
  }

  if (state.editingItemIndex !== null && state.inlineEditor) {
    state.inlineEditor.focus();
    if (!insertTextAtSelection(symbol)) {
      state.inlineEditor.innerText = `${state.inlineEditor.innerText || ""}${symbol}`;
    }
    state.inlineEditor.dispatchEvent(new Event("input", { bubbles: true }));
    return;
  }

  const input = elements.layer1Text;
  input.focus();
  const start = Number.isFinite(input.selectionStart) ? input.selectionStart : input.value.length;
  const end = Number.isFinite(input.selectionEnd) ? input.selectionEnd : input.value.length;
  input.setRangeText(symbol, start, end, "end");
  readFormIntoState();
  renderCanvasFromState();
  recordCanvasHistory();
  queueSave();
  queuePreview();
}

function makeDraggable(el, item, index) {
  el.addEventListener("mousedown", (e) => {
    if (item.locked) {
      return;
    }
    if (
      e.target.closest(".resize-handle") ||
      e.target.closest(".rotate-handle") ||
      e.target.closest(".delete-btn") ||
      e.target.closest(".edit-btn") ||
      e.target.closest(".canvas-inline-editor")
    ) {
      return;
    }

    e.preventDefault();
    state.isCanvasInteracting = true;
    state.selectedItem = index;
    elements.designCanvas.querySelectorAll(".canvas-item.selected").forEach((node) => {
      node.classList.remove("selected");
    });
    el.classList.add("selected");
    state.dragContext = {
      mode: "drag",
      item,
      index,
      el,
      offsetX: 0,
      offsetY: 0,
    };
    const scale = getCanvasPreviewScale();
    const itemRect = el.getBoundingClientRect();
    state.dragContext.offsetX = (e.clientX - itemRect.left) / scale;
    state.dragContext.offsetY = (e.clientY - itemRect.top) / scale;
    syncSelectedItemControls();
  });
}

function makeResizable(el, item, handle, index) {
  handle.addEventListener("mousedown", (e) => {
    if (item.locked) {
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    state.isCanvasInteracting = true;
    state.selectedItem = index;
    state.dragContext = {
      mode: "resize",
      item,
      index,
      el,
    };
    syncSelectedItemControls();
  });
}

function makeRotatable(el, item, handle, index) {
  handle.addEventListener("mousedown", (e) => {
    if (item.locked) {
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    state.isCanvasInteracting = true;
    state.selectedItem = index;
    state.dragContext = {
      mode: "rotate",
      item,
      index,
      el,
    };
    syncSelectedItemControls();
  });
}

function applySnapping(item, canvas, el) {
  const canvasWidth = canvas.clientWidth || 0;
  const canvasHeight = canvas.clientHeight || 0;
  const width = el.offsetWidth;
  const height = el.offsetHeight;
  const snapThreshold = 8;

  const current = {
    left: item.x,
    centerX: item.x + (width / 2),
    right: item.x + width,
    top: item.y,
    middleY: item.y + (height / 2),
    bottom: item.y + height,
  };

  const verticalCandidates = [{ position: canvasWidth / 2, anchor: "centerX" }];
  const horizontalCandidates = [{ position: canvasHeight / 2, anchor: "middleY" }];

  for (const other of state.canvasItems) {
    if (other === item) {
      continue;
    }
    verticalCandidates.push(
      { position: other.x, anchor: "left" },
      { position: other.x + (other.width / 2), anchor: "centerX" },
      { position: other.x + other.width, anchor: "right" }
    );
    horizontalCandidates.push(
      { position: other.y, anchor: "top" },
      { position: other.y + (other.height / 2), anchor: "middleY" },
      { position: other.y + other.height, anchor: "bottom" }
    );
  }

  let bestVertical = null;
  for (const candidate of verticalCandidates) {
    for (const anchor of ["left", "centerX", "right"]) {
      const distance = Math.abs(current[anchor] - candidate.position);
      if (distance > snapThreshold) {
        continue;
      }
      if (!bestVertical || distance < bestVertical.distance) {
        bestVertical = { ...candidate, anchor, distance };
      }
    }
  }

  let bestHorizontal = null;
  for (const candidate of horizontalCandidates) {
    for (const anchor of ["top", "middleY", "bottom"]) {
      const distance = Math.abs(current[anchor] - candidate.position);
      if (distance > snapThreshold) {
        continue;
      }
      if (!bestHorizontal || distance < bestHorizontal.distance) {
        bestHorizontal = { ...candidate, anchor, distance };
      }
    }
  }

  if (bestVertical) {
    if (bestVertical.anchor === "left") {
      item.x = bestVertical.position;
    } else if (bestVertical.anchor === "centerX") {
      item.x = bestVertical.position - (width / 2);
    } else {
      item.x = bestVertical.position - width;
    }
    elements.vGuide.style.left = `${bestVertical.position}px`;
    elements.vGuide.classList.remove("hidden");
  } else {
    elements.vGuide.classList.add("hidden");
  }

  if (bestHorizontal) {
    if (bestHorizontal.anchor === "top") {
      item.y = bestHorizontal.position;
    } else if (bestHorizontal.anchor === "middleY") {
      item.y = bestHorizontal.position - (height / 2);
    } else {
      item.y = bestHorizontal.position - height;
    }
    elements.hGuide.style.top = `${bestHorizontal.position}px`;
    elements.hGuide.classList.remove("hidden");
  } else {
    elements.hGuide.classList.add("hidden");
  }
}

function attachCanvasInteractions() {
  elements.designCanvas.addEventListener("click", () => {
    state.selectedItem = null;
    elements.vGuide.classList.add("hidden");
    elements.hGuide.classList.add("hidden");
    renderCanvasFromState();
  });

  window.addEventListener("mousemove", (e) => {
    if (!state.dragContext) {
      return;
    }

    const rect = elements.designCanvas.getBoundingClientRect();
    const scale = getCanvasPreviewScale();
    if (state.dragContext.mode === "resize") {
      let nextWidth = Math.max(40, ((e.clientX - rect.left) / scale) - state.dragContext.item.x);
      let nextHeight = Math.max(20, ((e.clientY - rect.top) / scale) - state.dragContext.item.y);
      const canvas = state.dragContext.el.parentElement;
      if (state.dragContext.item.type === "image") {
        const ratio =
          Number(state.dragContext.item.aspectRatio)
          || (Math.max(1, state.dragContext.item.width) / Math.max(1, state.dragContext.item.height));
        nextHeight = nextWidth / ratio;
      }
      const maxWidth = Math.max(40, (canvas.clientWidth || nextWidth) - state.dragContext.item.x);
      const maxHeight = Math.max(20, (canvas.clientHeight || nextHeight) - state.dragContext.item.y);
      const clampedWidth = Math.min(nextWidth, maxWidth);
      let clampedHeight = Math.min(nextHeight, maxHeight);
      if (state.dragContext.item.type === "image") {
        const ratio =
          Number(state.dragContext.item.aspectRatio)
          || (Math.max(1, state.dragContext.item.width) / Math.max(1, state.dragContext.item.height));
        clampedHeight = Math.max(20, Math.min(clampedWidth / ratio, maxHeight));
      }
      state.dragContext.item._tempWidth = clampedWidth;
      state.dragContext.item._tempHeight = clampedHeight;
      state.dragContext.el.style.width = `${clampedWidth}px`;
      state.dragContext.el.style.height = `${clampedHeight}px`;
    } else if (state.dragContext.mode === "rotate") {
      const elRect = state.dragContext.el.getBoundingClientRect();
      const cx = elRect.left + elRect.width / 2;
      const cy = elRect.top + elRect.height / 2;
      let angle = Math.atan2(e.clientY - cy, e.clientX - cx) * (180 / Math.PI);
      if (e.shiftKey) {
        const snapAngle = 45;
        angle = Math.round(angle / snapAngle) * snapAngle;
        state.dragContext.el.style.outline = "1px solid var(--accent)";
      } else {
        state.dragContext.el.style.outline = "";
      }
      state.dragContext.item.rotation = angle;
      state.dragContext.el.style.transform = `rotate(${angle}deg)`;
    } else {
      const nextX = ((e.clientX - rect.left) / scale) - state.dragContext.offsetX;
      const nextY = ((e.clientY - rect.top) / scale) - state.dragContext.offsetY;

      state.dragContext.item.x = nextX;
      state.dragContext.item.y = nextY;
      applyStickyBounds(state.dragContext.item, elements.designCanvas, state.dragContext.el);
      applySnapping(state.dragContext.item, elements.designCanvas, state.dragContext.el);
      state.dragContext.item._tempX = state.dragContext.item.x;
      state.dragContext.item._tempY = state.dragContext.item.y;
      state.dragContext.el.style.left = `${state.dragContext.item._tempX}px`;
      state.dragContext.el.style.top = `${state.dragContext.item._tempY}px`;
    }
  });

  window.addEventListener("mouseup", () => {
    if (state.dragContext) {
      if (state.dragContext.mode === "resize") {
        if (state.dragContext.item._tempWidth !== undefined) {
          state.dragContext.item.width = state.dragContext.item._tempWidth;
          state.dragContext.item.height = state.dragContext.item._tempHeight;
        }
        delete state.dragContext.item._tempWidth;
        delete state.dragContext.item._tempHeight;
        clampItemToCanvas(state.dragContext.item, elements.designCanvas, state.dragContext.el);
      } else if (state.dragContext.item._tempX !== undefined) {
        state.dragContext.item.x = state.dragContext.item._tempX;
        state.dragContext.item.y = state.dragContext.item._tempY;
        delete state.dragContext.item._tempX;
        delete state.dragContext.item._tempY;
        clampItemToCanvas(state.dragContext.item, elements.designCanvas, state.dragContext.el);
      } else if (state.dragContext.mode === "rotate") {
        state.dragContext.item.rotation = state.dragContext.item.rotation || 0;
        state.dragContext.el.style.outline = "";
      }
      syncCanvasIntoSettings();
      renderCanvasFromState();
      recordCanvasHistory();
      queueSave();
      triggerPreviewAfterInteraction();
    }
    state.dragContext = null;
    state.isCanvasInteracting = false;
    elements.canvasSafe.classList.remove("canvas-edge-hit");
    elements.vGuide.classList.add("hidden");
    elements.hGuide.classList.add("hidden");
  });
}

function startPrintProgress(total) {
  elements.printProgress.classList.remove("hidden");
  elements.printProgress.classList.add("active");
  elements.printProgress.textContent = `1 of ${total}`;
}

function updatePrintProgress(current, total) {
  elements.printProgress.textContent = `${current} of ${total}`;
}

function stopPrintProgress() {
  elements.printProgress.classList.remove("active");
  elements.printProgress.classList.add("hidden");
  elements.printProgress.textContent = "";
}

function applyTemplate(type) {
  switch (type) {
    case "text":
      state.canvasItems = [
        makeCanvasItem("text", {
          key: "layer1",
          text: state.settings.layer1.text || "Label",
          x: 40,
          y: 40,
          width: 280,
          height: 80,
        }),
      ];
      break;
    case "dual":
      state.canvasItems = [
        makeCanvasItem("text", {
          key: "layer1",
          text: state.settings.layer1.text || "Left",
          x: 28,
          y: 40,
          width: 160,
          height: 80,
        }),
        makeCanvasItem("text", {
          key: "layer2",
          text: state.settings.layer2.text || "Right",
          x: 200,
          y: 40,
          width: 160,
          height: 80,
        }),
      ];
      break;
    case "qr":
      state.canvasItems = [
        makeCanvasItem("text", {
          key: "layer1",
          text: state.settings.layer1.text || "Label",
          x: 26,
          y: 42,
          width: 180,
          height: 72,
        }),
        makeCanvasItem("qr", {
          key: "layer2",
          text: state.settings.layer2.text || "https://example.com",
          x: 230,
          y: 24,
          width: 120,
          height: 120,
        }),
      ];
      break;
    case "barcode":
      state.canvasItems = [
        makeCanvasItem("text", {
          key: "layer1",
          text: state.settings.layer1.text || "Label",
          x: 28,
          y: 20,
          width: 320,
          height: 44,
        }),
        makeCanvasItem("barcode", {
          key: "layer2",
          text: state.settings.layer2.text || "123456789",
          x: 28,
          y: 72,
          width: 320,
          height: 54,
        }),
      ];
      break;
    default:
      return;
  }

  log(`Template applied: ${type}`);
  state.selectedItem = 0;
  syncCanvasIntoSettings();
  renderCanvasFromState();
  recordCanvasHistory();
  queuePreview();
  queueSave();
}

function applySizePreset(size, button = null) {
  const [w, h] = size.split("x").map(Number);

  document.querySelectorAll("[data-size]").forEach((item) => {
    item.classList.remove("active");
  });
  if (button) {
    button.classList.add("active");
  }

  syncCanvasIntoSettings();
  state.settings.label_width_mm = w;
  state.settings.label_height_mm = h;
  state.settings.gap_mm = 5;

  const profile = state.profiles.find((item) =>
    Number(item.width_mm) === w && Number(item.height_mm) === h
  );

  if (profile) {
    state.settings.profile_id = profile.identifier;
  }

  state.canvasItems = [];
  updateCanvasSurface();
  syncForm();
  resetCanvasHistory();
  queuePreview();
  queueSave();
}

function applyCustomSize() {
  const w = Number(elements.customWidth.value);
  const h = Number(elements.customHeight.value);

  if (!w || !h) {
    showError("Enter valid width/height");
    return;
  }

  clearError();
  document.querySelectorAll("[data-size]").forEach((item) => {
    item.classList.remove("active");
  });
  state.settings.label_width_mm = w;
  state.settings.label_height_mm = h;
  state.canvasItems = [];
  updateCanvasSurface();
  syncForm();
  resetCanvasHistory();
  queuePreview();
  queueSave();
}

function applyExpiryLabel(dateValue = elements.expiryDate.value) {
  const formattedDate = formatExpiryDate(dateValue);
  if (!formattedDate) {
    showError("Pick an expiry date first");
    return;
  }

  clearError();
  state.canvasItems = buildExpiryCanvasItems(formattedDate);
  state.selectedItem = 1;
  syncCanvasIntoSettings();
  syncForm();
  recordCanvasHistory();
  queuePreview();
  queueSave();
}

function initElements() {
  Object.assign(elements, {
    appStatus: $("appStatus"),
    bleStatus: $("bleStatus"),
    connectionType: $("connectionType"),
    connectionStatus: $("connectionStatus"),
    connectionTimer: $("connectionTimer"),
    deviceConnectionBadge: $("deviceConnectionBadge"),
    deviceConnectionDetail: $("deviceConnectionDetail"),
    connectionTestBadge: $("connectionTestBadge"),
    batteryPanel: $("batteryPanel"),
    batteryMeter: $("batteryMeter"),
    batteryText: $("batteryText"),
    printProgress: $("printProgress"),
    errorBar: $("errorBar"),
    themeToggle: $("themeToggle"),
    modeToggle: $("modeToggle"),
    helpButton: $("helpButton"),
    onboardingOverlay: $("onboardingOverlay"),
    closeOnboarding: $("closeOnboarding"),
    helpOverlay: $("helpOverlay"),
    closeHelp: $("closeHelp"),
    recentLabels: $("recentLabels"),
    printHistory: $("printHistory"),
      expiryDate: $("expiryDate"),
      applyExpiryButton: $("applyExpiryButton"),
      layer1Text: $("layer1Text"),
      layer1Mode: $("layer1Mode"),
      layer1Align: $("layer1Align"),
      invertToggle: $("invertToggle"),
      imageMode: $("imageMode"),
      autoImage: $("autoImage"),
      edgeEnhance: $("edgeEnhance"),
      brightness: $("brightness"),
      contrast: $("contrast"),
      threshold: $("threshold"),
      fontPicker: $("fontPicker"),
      fontSizeInput: $("fontSizeInput"),
      systemFontToggle: $("systemFontToggle"),
      toggleBoldButton: $("toggleBoldButton"),
      toggleUppercaseButton: $("toggleUppercaseButton"),
      addTextButton: $("addTextButton"),
      addQrButton: $("addQrButton"),
      addBarcodeButton: $("addBarcodeButton"),
      addImageButton: $("addImageButton"),
      undoButton: $("undoButton"),
      redoButton: $("redoButton"),
      duplicateButton: $("duplicateButton"),
      lockButton: $("lockButton"),
      sendBackwardButton: $("sendBackwardButton"),
      bringForwardButton: $("bringForwardButton"),
      fitTextButton: $("fitTextButton"),
      exportPngButton: $("exportPngButton"),
      imageUpload: $("imageUpload"),
      saveDesignButton: $("saveDesignButton"),
      loadDesignButton: $("loadDesignButton"),
      savedDesigns: $("savedDesigns"),
      customWidth: $("customWidth"),
      customHeight: $("customHeight"),
      applyCustomSize: $("applyCustomSize"),
    copies: $("copies"),
    copiesPlus: $("copiesPlus"),
    copiesMinus: $("copiesMinus"),
    portSelect: $("portSelect"),
    connectPortButton: $("connectPortButton"),
    refreshPortsButton: $("refreshPortsButton"),
    bleDeviceSelect: $("bleDeviceSelect"),
    scanBleButton: $("scanBleButton"),
    connectBleButton: $("connectBleButton"),
    disconnectBleButton: $("disconnectBleButton"),
    testPrintButton: $("testPrintButton"),
    stopPrintButton: $("stopPrintButton"),
    printButton: $("printButton"),
    designCanvasViewport: $("designCanvasViewport"),
    designCanvas: $("designCanvas"),
    canvasSafe: $("canvasSafe"),
    vGuide: $("vGuide"),
    hGuide: $("hGuide"),
    printPreview: $("printPreview"),
  });
}

function log(message) {
  console.log(`[pigeon] ${message}`);
}

function setStatus(text) {
  elements.appStatus.textContent = text;
}

function setBusy(busy, text) {
  state.busy = busy;
  if (text) {
    setStatus(text);
  }
  const controls = [
    elements.themeToggle,
    elements.connectPortButton,
    elements.refreshPortsButton,
    elements.scanBleButton,
    elements.connectBleButton,
    elements.disconnectBleButton,
    elements.testPrintButton,
    elements.printButton,
  ];
  for (const control of controls) {
    control.disabled = busy;
  }
  updateStopPrintButton();
}

function updateStopPrintButton() {
  if (!elements.stopPrintButton) {
    return;
  }
  elements.stopPrintButton.disabled = !state.printQueueActive;
  elements.stopPrintButton.textContent = state.printStopRequested ? "Stopping..." : "Stop";
}

function requestStopPrintQueue() {
  if (!state.busy) {
    return;
  }
  state.printStopRequested = true;
  updateStopPrintButton();
  setStatus("Stopping queue...");
}

function updateConnectionUI() {
  const isBLE = state.settings.output_mode === "BLE";

  elements.connectionType.textContent = isBLE ? "BLE" : "COM";

  let statusText = "";
  let statusClass = "";
  let detailText = "";

  if (isBLE) {
    if (state.connecting) {
      statusText = "Connecting...";
      statusClass = "status-connecting";
      detailText = `Connecting to ${state.settings.ble_device_name || state.settings.ble_device_address || "BLE device"}`;
    } else if (state.bleState.connected) {
      statusText = `Connected (${state.bleState.name || state.bleState.address})`;
      statusClass = "status-connected";
      detailText = `BLE connected to ${state.bleState.name || state.bleState.address}`;
    } else {
      statusText = "Disconnected";
      statusClass = "status-disconnected";
      detailText = state.settings.ble_device_name || state.settings.ble_device_address
        ? `Ready to connect to ${state.settings.ble_device_name || state.settings.ble_device_address}`
        : "No BLE device connected";
    }
  } else {
    if (!state.settings.port) {
      statusText = "Disconnected";
      statusClass = "status-disconnected";
      detailText = "No COM port selected";
    } else if (state.comConnecting) {
      statusText = "Connecting...";
      statusClass = "status-connecting";
      detailText = `Opening COM printer on ${state.settings.port}`;
    } else if (state.comConnected) {
      statusText = `Connected (${state.settings.port})`;
      statusClass = "status-connected";
      detailText = `COM printer connected on ${state.settings.port}`;
    } else {
      statusText = `Selected (${state.settings.port})`;
      statusClass = "status-connecting";
      detailText = `COM port ${state.settings.port} selected. Click Connect COM to open it.`;
    }
  }

  elements.connectionStatus.textContent = statusText;
  elements.connectionStatus.className = `status-indicator ${statusClass}`;
  if (elements.deviceConnectionBadge) {
    elements.deviceConnectionBadge.textContent = statusText;
    elements.deviceConnectionBadge.className = `status-pill${statusClass === "status-disconnected" ? " muted" : ""}`;
  }
  if (elements.deviceConnectionDetail) {
    elements.deviceConnectionDetail.textContent = detailText;
  }
}

function updateBatteryUI() {
  const level = Number(state.batteryLevel);
  const showBattery = state.settings?.output_mode === "BLE" && state.bleState.connected;
  const hasBattery = showBattery && Number.isFinite(level) && level > 0;

  if (!showBattery) {
    elements.batteryPanel.classList.add("hidden");
    elements.batteryMeter.innerHTML = "";
    elements.batteryText.textContent = "--%";
    return;
  }

  if (!hasBattery) {
    elements.batteryPanel.classList.remove("hidden");
    elements.batteryMeter.innerHTML = "";
    for (let index = 0; index < 5; index += 1) {
      const block = document.createElement("span");
      block.className = "battery-block";
      elements.batteryMeter.appendChild(block);
    }
    elements.batteryText.textContent = "Unknown";
    return;
  }

  const clamped = Math.max(0, Math.min(100, Math.round(level)));
  const filledBlocks = Math.max(1, Math.ceil(clamped / 20));
  elements.batteryPanel.classList.remove("hidden");
  elements.batteryMeter.innerHTML = "";

  for (let index = 0; index < 5; index += 1) {
    const block = document.createElement("span");
    block.className = "battery-block";
    if (index < filledBlocks) {
      block.classList.add("active");
      if (clamped <= 20) {
        block.classList.add("low");
      } else if (clamped <= 40) {
        block.classList.add("warn");
      }
    }
    elements.batteryMeter.appendChild(block);
  }

  elements.batteryText.textContent = `${clamped}%`;
}

async function refreshBleBattery() {
  const address = state.bleState.address || state.settings.ble_device_address;
  if (!state.bleState.connected || !address) {
    state.batteryLevel = null;
    updateBatteryUI();
    return;
  }

  try {
    const result = await safeRequest("bleBattery", {
      address,
      pair: false,
    });
    state.batteryLevel = Number.isFinite(Number(result.battery)) ? Number(result.battery) : null;
  } catch (_error) {
    state.batteryLevel = null;
  }

  updateBatteryUI();
}

function startConnectionTimer() {
  state.connectionStartTime = Date.now();

  clearInterval(state.connectionTimerInterval);

  state.connectionTimerInterval = setInterval(() => {
    const elapsed = Date.now() - state.connectionStartTime;

    const seconds = Math.floor(elapsed / 1000);
    const mins = String(Math.floor(seconds / 60)).padStart(2, "0");
    const secs = String(seconds % 60).padStart(2, "0");

    elements.connectionTimer.textContent = `${mins}:${secs}`;
  }, 1000);

  elements.connectionTimer.classList.remove("hidden");
}

function stopConnectionTimer() {
  clearInterval(state.connectionTimerInterval);
  state.connectionTimerInterval = null;
  state.connectionStartTime = null;
  elements.connectionTimer.classList.add("hidden");
  elements.connectionTimer.textContent = "00:00";
}

function fillSelect(select, items, getValue, getLabel) {
  select.innerHTML = "";
  for (const item of items) {
    const option = document.createElement("option");
    option.value = getValue(item);
    option.textContent = getLabel(item);
    select.appendChild(option);
  }
}

async function runComConnect() {
  try {
    clearError();
    readFormIntoState();
    if (!state.settings.port) {
      showError("Select a COM port first");
      return;
    }

    state.comConnecting = true;
    updateConnectionUI();
    setBusy(true, "Connecting COM");
    const result = await safeRequest("connectSerial", {
      settings: state.settings,
    });
    state.settings.output_mode = "Printer";
    state.comConnecting = false;
    state.comConnected = true;
    state.batteryLevel = null;
    updateBatteryUI();
    updateConnectionTestBadge("idle");
      startConnectionTimer();
      queueSave();
      updateConnectionUI();
    setStatus(`COM connected to ${result.port}`);
  } catch (error) {
    state.comConnecting = false;
    state.comConnected = false;
    state.batteryLevel = null;
    updateBatteryUI();
      updateConnectionTestBadge("fail");
      stopConnectionTimer();
      updateConnectionUI();
      showError(error.message);
  } finally {
    state.comConnecting = false;
    setBusy(false);
  }
}

function prioritizeDevices(list, lastUsed) {
  if (!lastUsed) {
    return list;
  }

  const index = list.findIndex((item) => item.address === lastUsed || item === lastUsed);
  if (index === -1) {
    return list;
  }

  const selected = list.splice(index, 1)[0];
  return [selected, ...list];
}

function syncStaticChoices(initData) {
  fillSelect(elements.layer1Mode, initData.layerModes, (item) => item, (item) => item);
  fillSelect(elements.layer1Align, initData.alignments, (item) => item, (item) => item);
}

function syncPorts() {
  const sortedPorts = prioritizeDevices([...state.serialPorts], state.settings.port);
  fillSelect(elements.portSelect, sortedPorts, (item) => item, (item) => item);
  if (state.settings.port && sortedPorts.includes(state.settings.port)) {
    elements.portSelect.value = state.settings.port;
  } else {
    elements.portSelect.value = "";
    state.settings.port = "";
    state.comConnected = false;
  }
}

function syncBleDevices() {
  const sorted = prioritizeDevices([...state.bleDevices], state.settings.ble_device_address);
  fillSelect(
    elements.bleDeviceSelect,
    sorted,
    (item) => item.address,
    (item) => item.label || `${item.name} [${item.address}]`
  );
  if (state.settings.ble_device_address) {
    elements.bleDeviceSelect.value = state.settings.ble_device_address;
  } else if (sorted.length > 0) {
    elements.bleDeviceSelect.value = sorted[0].address;
    state.settings.ble_device_address = sorted[0].address;
    state.settings.ble_device_name = sorted[0].name || "";
  }
}

function syncForm() {
  state.isHydrating = true;
  const selected = state.selectedItem !== null ? state.canvasItems[state.selectedItem] : state.canvasItems[0];
  elements.layer1Text.value = selected?.text || "";
  elements.layer1Mode.value = state.settings.layer1.mode;
  elements.layer1Align.value = state.settings.layer1.align;
  elements.invertToggle.checked = Boolean(state.settings.invert);
  elements.imageMode.value = state.settings.image_mode || "threshold";
  elements.autoImage.checked = Boolean(state.settings.auto_image);
  elements.edgeEnhance.checked = Boolean(state.settings.edge_enhance);
  elements.brightness.value = String(state.settings.brightness ?? 1);
  elements.contrast.value = String(state.settings.contrast ?? 2);
  elements.threshold.value = String(state.settings.threshold ?? 180);
  elements.systemFontToggle.checked = state.useSystemFonts;
  elements.copies.value = state.settings.copies;
  elements.customWidth.value = String(state.settings.label_width_mm ?? "");
  elements.customHeight.value = String(state.settings.label_height_mm ?? "");
  if (elements.expiryDate && !elements.expiryDate.value) {
    elements.expiryDate.value = "";
  }
  syncPorts();
    syncBleDevices();
    refreshBleStatus();
    updateConnectionUI();
    updateConnectionTestBadge();
    updateBatteryUI();
    updateCanvasSurface();
  updatePrintPreviewOrientation();
  buildCanvasItemsFromSettings();
  renderCanvasFromState();
  renderSavedDesigns();
  state.isHydrating = false;
}

function refreshBleStatus() {
  const stateText = state.bleState.connected
    ? `BLE connected to ${state.bleState.name || state.bleState.address}`
    : "BLE disconnected";
  elements.bleStatus.textContent = stateText;
}

function forceSimpleDefaults() {
  if (!state.settings.layer1.mode) {
    state.settings.layer1.mode = "Text";
  }
  if (!state.settings.layer1.align) {
    state.settings.layer1.align = "Center";
  }
  if (!state.settings.layer2.mode) {
    state.settings.layer2.mode = "Off";
  }
  if (!state.settings.layer2.align) {
    state.settings.layer2.align = "Right";
  }
  if (!state.settings.output_mode) {
    state.settings.output_mode = "Printer";
  }
  state.settings.baud_rate = 115200;
  state.settings.invert = false;
}

function readFormIntoState() {
  const selectedItem = state.selectedItem !== null ? state.canvasItems[state.selectedItem] : state.canvasItems[0];
  if (selectedItem) {
    if (selectedItem.type !== "image") {
      selectedItem.text = elements.layer1Text.value;
    }
    selectedItem.invert = Boolean(elements.invertToggle.checked);
    if (selectedItem.type === "text") {
      selectedItem.font = elements.fontPicker.value || selectedItem.font || state.settings.font_name;
    }
  }
  state.settings.invert = Boolean(elements.invertToggle.checked);
  state.settings.image_mode = elements.imageMode.value || "threshold";
  state.settings.auto_image = Boolean(elements.autoImage.checked);
  state.settings.edge_enhance = Boolean(elements.edgeEnhance.checked);
  state.settings.brightness = Number(elements.brightness.value) || 1;
  state.settings.contrast = Number(elements.contrast.value) || 2;
  state.settings.threshold = Number(elements.threshold.value) || 180;
  state.settings.layer1.mode = elements.layer1Mode.value || "Text";
  state.settings.layer1.align = elements.layer1Align.value;
  state.settings.copies = Math.max(1, Number(elements.copies.value) || 1);
  if (state.settings.port !== elements.portSelect.value) {
    state.comConnected = false;
    state.comConnecting = false;
  }
  state.settings.port = elements.portSelect.value;
  state.settings.baud_rate = 115200;
  state.settings.ble_device_address = elements.bleDeviceSelect.value;
  state.settings.ble_device_name =
      state.bleDevices.find((item) => item.address === elements.bleDeviceSelect.value)?.name || "";
  state.settings.ble_write_char_uuid = state.bleWritable[0]?.uuid || "";
  state.settings.output_mode = state.bleState.connected && !state.comConnected ? "BLE" : "Printer";
  syncCanvasIntoSettings();
}

async function request(command, params = {}) {
  return window.pigeonApi.request(command, params);
}

function showBackendRestart() {
  setStatus("Reconnecting backend...");
}

function isBackendRestartableError(error) {
  const message = String(error?.message || "");
  return message.includes("Backend crashed") || message.includes("Backend exited");
}

async function safeRequest(command, params = {}) {
  try {
    return await request(command, params);
  } catch (error) {
    if (!isBackendRestartableError(error)) {
      throw error;
    }

    showBackendRestart();
    log("Backend restarted, retrying...");
    await new Promise((resolve) => setTimeout(resolve, 1200));
    return await request(command, params);
  }
}

function updateImageSettings() {
  if (state.isHydrating) {
    return;
  }
  clearError();
  state.settings.image_mode = elements.imageMode.value;
  state.settings.brightness = Number(elements.brightness.value);
  state.settings.contrast = Number(elements.contrast.value);
  state.settings.threshold = Number(elements.threshold.value);
  state.settings.invert = elements.invertToggle.checked;
  queuePreview();
  queueSave();
}

function updateAdvancedImageSettings() {
  if (state.isHydrating) {
    return;
  }
  clearError();
  state.settings.auto_image = elements.autoImage.checked;
  state.settings.edge_enhance = elements.edgeEnhance.checked;
  queuePreview();
  queueSave();
}

function adjustCopies(delta) {
  let value = Number(elements.copies.value) || 1;
  value += delta;

  if (value < 1) {
    value = 1;
  }

  elements.copies.value = value;
  state.settings.copies = value;
  clearError();
  queuePreview();
  queueSave();
}

function queueSave() {
  clearTimeout(state.saveTimer);
  state.saveTimer = setTimeout(async () => {
    try {
      saveDesign();
      await safeRequest("saveSettings", { settings: state.settings });
    } catch (error) {
      log(`Save settings failed: ${error.message}`);
    }
  }, 450);
}

function queuePreview() {
  clearTimeout(state.previewTimer);
  state.previewTimer = setTimeout(() => {
    updatePreview();
  }, 80);
}

function triggerPreviewAfterInteraction() {
  clearTimeout(state.previewTimer);

  state.previewTimer = setTimeout(() => {
    updatePreview();
  }, 100);
}

async function updatePreview(options = {}) {
  const { force = false } = options;
  if (state.busy && !force) {
    return;
  }
  const isInlineEditing = state.editingItemIndex !== null && Boolean(state.inlineEditor);
  try {
    if (!isInlineEditing) {
      setStatus("Updating...");
      elements.designCanvas.style.opacity = "0.6";
      elements.printPreview.style.opacity = "0.6";
    }
    if (state.editingItemIndex !== null && state.inlineEditor) {
      const editingItem = state.canvasItems[state.editingItemIndex];
      if (editingItem && editingItem.type !== "image") {
        const liveText = normalizeInlineEditorText(state.inlineEditor.innerText);
        editingItem.text = liveText;
        if (state.selectedItem === state.editingItemIndex) {
          elements.layer1Text.value = liveText;
        }
        syncCanvasIntoSettings();
      }
      updateCanvasSurface();
      updatePrintPreviewOrientation();
    } else {
      readFormIntoState();
      updateCanvasSurface();
      updatePrintPreviewOrientation();
      buildCanvasItemsFromSettings();
      renderCanvasFromState();
    }
    const imageData = await captureCanvasImage();
    const result = await safeRequest("preview", {
      settings: buildRasterRenderSettings(),
      imagePath: imageData,
    });
    elements.printPreview.src = result.printImage;
    elements.designCanvas.style.opacity = "1";
    elements.printPreview.style.opacity = "1";
    if (!isInlineEditing) {
      setStatus("Preview updated");
    }
  } catch (error) {
    elements.designCanvas.style.opacity = "1";
    elements.printPreview.style.opacity = "1";
    log(`Preview failed: ${error.message}`);
    if (!isInlineEditing) {
      setStatus("Preview failed");
    }
  }
}

async function refreshPorts() {
  const result = await safeRequest("listSerialPorts");
  state.serialPorts = result.ports;
  syncPorts();
  if (state.settings.port && state.serialPorts.includes(state.settings.port)) {
    elements.portSelect.value = state.settings.port;
  } else {
    elements.portSelect.value = "";
    state.settings.port = "";
    state.comConnected = false;
  }
  if (elements.portSelect.value) {
    state.settings.port = elements.portSelect.value;
    queueSave();
  }
  if (state.settings.port) {
    log(`Using COM port: ${state.settings.port}`);
  }
  if (state.settings.output_mode !== "BLE" && state.settings.port && state.comConnected) {
    startConnectionTimer();
  } else if (state.settings.output_mode !== "BLE") {
    stopConnectionTimer();
  }
  updateConnectionUI();
  log(`Detected ${state.serialPorts.length} serial port(s).`);
}

async function refreshBleState() {
  const result = await safeRequest("bleState");
  state.bleState = result.state;
  refreshBleStatus();
  updateConnectionUI();
  await refreshBleBattery();
}

async function runTestPrint() {
  try {
    clearError();
    readFormIntoState();
    setBusy(true, "Sending test print");

    await safeRequest("testPrint", {
      settings: state.settings,
    });

    log("Test print sent");
    updateConnectionTestBadge("pass");
    setStatus("Test print complete");
  } catch (error) {
    updateConnectionTestBadge("fail");
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

async function runPrint() {
  try {
    clearError();
    readFormIntoState();
    const error = validateBeforePrint();
    if (error) {
      showError(error);
      return;
    }

    const total = state.settings.copies || 1;
    state.printStopRequested = false;

    state.printQueueActive = true;
    setBusy(true, "Printing...");
    startPrintProgress(total);
    renderCanvasFromState();
    const imageData = await captureCanvasImage();
    const rasterSettings = buildRasterRenderSettings();

    let lastResult = null;
    for (let i = 1; i <= total; i += 1) {
      if (state.printStopRequested) {
        break;
      }
      lastResult = await safeRequest("print", {
        settings: { ...rasterSettings, copies: 1 },
        imagePath: imageData,
      });
      updatePrintProgress(i, total);
    }

    if (lastResult) {
      await refreshBleState();
      updateConnectionTestBadge("pass");
      if (lastResult?.mode === "BLE") {
        log(
          `BLE print sent: protocol=${lastResult.result.protocol}, bytes=${lastResult.result.bytes_sent}, packets=${lastResult.result.packet_count}, notifications=${lastResult.result.notification_count}`
        );
      } else if (lastResult?.mode === "Printer") {
        log(`Serial print sent to ${lastResult.port}`);
      }
    }

    stopPrintProgress();
    if (state.printStopRequested) {
      addPrintHistory(false, "Print queue stopped");
      setStatus("Print queue stopped");
      log("Print queue stopped by user");
    } else {
      saveRecentLabel(state.canvasItems.find((item) => item.type === "text")?.text || "");
      addPrintHistory(true, "Print successful");
      setStatus("Printed successfully");
      showPrintSuccess();
    }
  } catch (error) {
    log(`Print failed: ${error.message}`);
    updateConnectionTestBadge("fail");
    stopPrintProgress();
    addPrintHistory(false, "Print failed");
    setStatus("Print failed");
    showError(error.message);
  } finally {
    state.printQueueActive = false;
    state.printStopRequested = false;
    stopPrintProgress();
    setBusy(false);
  }
}

async function runBleScan() {
  try {
    setBusy(true, "Scanning BLE");
    const result = await safeRequest("scanBle", { timeout: state.settings.ble_scan_timeout || 5.0 });
    state.bleDevices = result.devices;
    syncBleDevices();
    setStatus(`Found ${state.bleDevices.length} BLE device(s)`);
  } catch (error) {
    log(`BLE scan failed: ${error.message}`);
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

async function runBleConnect(options = {}) {
  const { silent = false } = options;
  try {
    readFormIntoState();
    if (!state.settings.ble_device_address) {
      if (!silent) {
        showError("Select a BLE device first.");
      }
      return;
    }
    clearError();
    state.connecting = true;
    setBusy(true, "Connecting BLE");
    updateConnectionUI();
    const inspect = await safeRequest("inspectBle", {
      address: state.settings.ble_device_address,
      pair: false,
    });
    state.bleWritable = inspect.writable;
    state.bleNotify = inspect.notify;
    state.settings.ble_write_char_uuid =
      state.bleWritable.find((item) => item.preferred)?.uuid
      || state.bleWritable[0]?.uuid
      || "";
    state.settings.ble_write_with_response = false;
    state.settings.ble_pair = false;
    state.settings.ble_chunk_size = 180;
    log(`BLE ready (${state.bleWritable.length} endpoints found)`);
    const result = await safeRequest("connectBle", {
      address: state.settings.ble_device_address,
      pair: false,
    });
    state.bleState = result.state;
    state.settings.ble_device_address = elements.bleDeviceSelect.value;
    state.settings.ble_device_name =
      state.bleDevices.find((item) => item.address === elements.bleDeviceSelect.value)?.name
      || state.bleState.name
      || "";
    state.settings.output_mode = "BLE";
    state.comConnected = false;
    state.connecting = false;
    setBusy(false);
    syncForm();
    updateConnectionUI();
    startConnectionTimer();
    await refreshBleBattery();
    updateConnectionTestBadge("idle");
    queueSave();
    queuePreview();
    setStatus(`BLE connected to ${state.bleState.name || state.bleState.address}`);
  } catch (error) {
    state.connecting = false;
    updateConnectionTestBadge("fail");
    log(`BLE connect failed: ${error.message}`);
    updateConnectionUI();
    if (!silent) {
      showError(error.message);
    }
    throw error;
  } finally {
    state.connecting = false;
    setBusy(false);
    updateConnectionUI();
  }
}

async function runBleDisconnect() {
  try {
    setBusy(true, "Disconnecting BLE");
    const result = await safeRequest("disconnectBle");
      state.bleState = result.state;
      state.bleWritable = [];
      state.bleNotify = [];
      state.batteryLevel = null;
      updateBatteryUI();
      updateConnectionTestBadge("idle");
      state.settings.ble_write_char_uuid = "";
      state.settings.output_mode = "Printer";
    refreshBleStatus();
    updateConnectionUI();
    stopConnectionTimer();
    queueSave();
    setStatus("BLE disconnected");
  } catch (error) {
    log(`BLE disconnect failed: ${error.message}`);
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

function attachFormListeners() {
  const inputs = [
    elements.layer1Mode,
    elements.layer1Align,
    elements.copies,
    elements.portSelect,
    elements.bleDeviceSelect,
  ];

  for (const input of inputs) {
    input.addEventListener("input", () => {
      if (state.isHydrating) {
        return;
      }
      clearError();
      readFormIntoState();
      queueSave();
      queuePreview();
    });
    input.addEventListener("change", () => {
      if (state.isHydrating) {
        return;
      }
      clearError();
      readFormIntoState();
      queueSave();
      queuePreview();
    });
  }

  elements.layer1Text.addEventListener("input", () => {
    if (state.isHydrating) {
      return;
    }
    clearError();
    readFormIntoState();
    queueSave();
    clearTimeout(state.previewTimer);
    state.previewTimer = setTimeout(updatePreview, 40);
  });

  elements.layer1Text.addEventListener("change", () => {
      if (state.isHydrating) {
        return;
      }
      clearError();
      readFormIntoState();
      recordCanvasHistory();
      queueSave();
      queuePreview();
    });

    elements.fontPicker.addEventListener("change", () => {
      if (state.selectedItem === null) {
        return;
      }
      state.canvasItems[state.selectedItem].font = elements.fontPicker.value;
      clearError();
      renderCanvasFromState();
      recordCanvasHistory();
      queueSave();
      queuePreview();
    });

    elements.fontSizeInput.addEventListener("input", () => {
      if (state.selectedItem === null) {
        return;
      }
      const value = String(elements.fontSizeInput.value ?? "").trim();
      if (!value) {
        return;
      }

      const numeric = Number(value);
      if (!Number.isFinite(numeric) || numeric < 6) {
        return;
      }

      commitSelectedFontSize(value);
    });

    elements.fontSizeInput.addEventListener("change", () => {
      commitSelectedFontSize(elements.fontSizeInput.value, { restoreIfEmpty: true });
    });

    elements.fontSizeInput.addEventListener("blur", () => {
      commitSelectedFontSize(elements.fontSizeInput.value, { restoreIfEmpty: true });
    });

    elements.systemFontToggle.addEventListener("change", async () => {
      state.useSystemFonts = elements.systemFontToggle.checked;
      localStorage.setItem("use_system_fonts", state.useSystemFonts ? "true" : "false");
      await loadFonts();
      queueSave();
    });

    elements.toggleBoldButton.addEventListener("click", () => {
      if (state.selectedItem === null) {
        return;
      }
      const item = state.canvasItems[state.selectedItem];
      if (item.type !== "text") {
        return;
      }
      item.fontWeight = Number(item.fontWeight) >= 700 ? 400 : 700;
      syncCanvasIntoSettings();
      renderCanvasFromState();
      recordCanvasHistory();
      queueSave();
      queuePreview();
    });

    elements.toggleUppercaseButton.addEventListener("click", () => {
      if (state.selectedItem === null) {
        return;
      }
      const item = state.canvasItems[state.selectedItem];
      if (item.type !== "text") {
        return;
      }
      item.textTransform = item.textTransform === "uppercase" ? "none" : "uppercase";
      syncCanvasIntoSettings();
      renderCanvasFromState();
      recordCanvasHistory();
      queueSave();
      queuePreview();
    });

    elements.imageMode.addEventListener("change", updateImageSettings);
    elements.brightness.addEventListener("input", updateImageSettings);
    elements.contrast.addEventListener("input", updateImageSettings);
    elements.threshold.addEventListener("input", updateImageSettings);
    elements.invertToggle.addEventListener("change", updateImageSettings);
    elements.autoImage.addEventListener("change", updateAdvancedImageSettings);
    elements.edgeEnhance.addEventListener("change", updateAdvancedImageSettings);

    elements.themeToggle.addEventListener("click", toggleTheme);
  elements.modeToggle.addEventListener("click", toggleMode);
  elements.helpButton.addEventListener("click", openHelp);
  elements.closeOnboarding.addEventListener("click", closeOnboarding);
  elements.closeHelp.addEventListener("click", closeHelp);
  elements.addTextButton.addEventListener("click", () => addItem("text"));
  elements.addQrButton.addEventListener("click", () => addItem("qr"));
  elements.addBarcodeButton.addEventListener("click", () => addItem("barcode"));
  elements.addImageButton.addEventListener("click", () => elements.imageUpload.click());
  elements.imageUpload.addEventListener("change", (e) => {
    const [file] = e.target.files || [];
    if (!file) {
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      addItem("image", {
        text: file.name,
        imageData: String(reader.result || ""),
        width: 120,
        height: 120,
      });
      elements.imageUpload.value = "";
    };
    reader.readAsDataURL(file);
  });
  elements.saveDesignButton.addEventListener("click", async () => {
    readFormIntoState();
    await saveDesign();
    setStatus("Design saved");
  });
  elements.loadDesignButton.addEventListener("click", () => {
    if (loadDesign()) {
      queuePreview();
      queueSave();
      setStatus("Design loaded");
    }
  });
  elements.undoButton.addEventListener("click", undoCanvas);
  elements.redoButton.addEventListener("click", redoCanvas);
  elements.duplicateButton.addEventListener("click", duplicateSelectedItem);
  elements.lockButton.addEventListener("click", toggleSelectedLock);
  elements.sendBackwardButton.addEventListener("click", () => moveSelectedLayer("backward"));
  elements.bringForwardButton.addEventListener("click", () => moveSelectedLayer("forward"));
  elements.fitTextButton.addEventListener("click", fitSelectedTextToBox);
  elements.exportPngButton.addEventListener("click", exportCurrentDesignPng);
  elements.copiesPlus.addEventListener("click", () => adjustCopies(1));
  elements.copiesMinus.addEventListener("click", () => adjustCopies(-1));
  elements.applyCustomSize.addEventListener("click", applyCustomSize);
  elements.applyExpiryButton.addEventListener("click", () => applyExpiryLabel());
  elements.expiryDate.addEventListener("change", () => {
    clearError();
    applyExpiryLabel(elements.expiryDate.value);
  });
  elements.connectPortButton.addEventListener("click", runComConnect);
  elements.refreshPortsButton.addEventListener("click", refreshPorts);
  elements.scanBleButton.addEventListener("click", runBleScan);
  elements.connectBleButton.addEventListener("click", runBleConnect);
  elements.disconnectBleButton.addEventListener("click", runBleDisconnect);
  elements.testPrintButton.addEventListener("click", runTestPrint);
  elements.stopPrintButton.addEventListener("click", requestStopPrintQueue);
  elements.printButton.addEventListener("click", runPrint);

  document.querySelectorAll("[data-size]").forEach((button) => {
    button.addEventListener("click", () => {
      applySizePreset(button.dataset.size, button);
    });
  });

  document.querySelectorAll("[data-template]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      const type = e.currentTarget.dataset.template;
      applyTemplate(type);
    });
  });

  document.querySelectorAll("[data-align]").forEach((btn) => {
    btn.addEventListener("click", () => {
      alignSelected(btn.dataset.align);
    });
  });

  document.querySelectorAll("[data-symbol]").forEach((button) => {
    button.addEventListener("click", () => {
      insertSymbol(button.dataset.symbol || "");
    });
  });

  elements.layer1Mode.addEventListener("change", () => {
    if (state.isHydrating || state.selectedItem === null) {
      return;
    }
    const item = state.canvasItems[state.selectedItem];
    const nextMode = elements.layer1Mode.value;
    item.type = nextMode === "QR" ? "qr" : nextMode === "Barcode" ? "barcode" : "text";
    if (item.type === "text" && !item.text) {
      item.text = "New Text";
    }
    if (item.type === "qr" && !item.text) {
      item.text = "QR DATA";
    }
    if (item.type === "barcode" && !item.text) {
      item.text = "123456789";
    }
    syncCanvasIntoSettings();
    renderCanvasFromState();
    recordCanvasHistory();
    queueSave();
    queuePreview();
  });
}

async function boot() {
    loadTheme();
    initElements();
    initSidebarCollapsibles();
    attachCanvasInteractions();
    loadMode();
    state.useSystemFonts = localStorage.getItem("use_system_fonts") === "true";
    checkOnboarding();
    attachFormListeners();
    setupShortcuts();
    localStorage.removeItem("print_history");
    const initData = await safeRequest("init");
    state.settings = structuredClone(initData.settings);
    state.profiles = initData.profiles || [];
    state.builtInFonts = initData.fonts || [];
    state.settings.output_mode = "Printer";
    state.settings.port = "";
    state.settings.baud_rate = 115200;
    state.settings.invert = false;
    state.settings.ble_device_address = "";
    state.settings.ble_device_name = "";
    state.bleState = { connected: false, address: "", name: "" };
    state.comConnected = false;
    syncStaticChoices(initData);
    forceSimpleDefaults();
    await loadFonts();
    loadSavedDesigns();
    renderRecentLabels();
    renderPrintHistory();
    await refreshPorts();
    await refreshBleState();
    stopConnectionTimer();
    const startupCanvas = getCanvasLogicalSize();
    const startupWidth = 200;
    const startupHeight = 60;
    state.canvasItems = [
      makeCanvasItem("text", {
        id: 1,
        key: "layer1",
        text: "Label",
        x: Math.max(0, Math.round((startupCanvas.width - startupWidth) / 2)),
        y: Math.max(0, Math.round((startupCanvas.height - startupHeight) / 2)),
        width: startupWidth,
        height: startupHeight,
      }),
    ];
    state.selectedItem = 0;
    syncCanvasIntoSettings();
    updateCanvasSurface();
    updatePrintPreviewOrientation();
    syncForm();
    renderCanvasFromState();
    resetCanvasHistory();
    updateCanvasActionButtons();
    await updatePreview();
    setStatus("Ready");
  }

boot().catch((error) => {
  console.error(error);
  showError(error.message);
});
