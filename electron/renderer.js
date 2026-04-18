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
  isHydrating: false,
  previewTimer: null,
  saveTimer: null,
  busy: false,
  connecting: false,
  connectionStartTime: null,
  connectionTimerInterval: null,
  isCanvasInteracting: false,
  canvasItems: [],
  selectedItem: null,
  dragContext: null,
  canvasBackground: "",
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

function showError(message) {
  elements.errorBar.textContent = message;
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

    if (!isTypingTarget && e.key === "Delete" && state.selectedItem !== null) {
      e.preventDefault();
      deleteItem(state.selectedItem);
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
    rotation: 0,
    invert: false,
    imageData: "",
    aspectRatio: type === "image" ? 1 : null,
  };
  const item = { ...defaults, ...overrides };
  if (item.type === "image") {
    const width = Math.max(1, Number(item.width) || defaults.width);
    const height = Math.max(1, Number(item.height) || defaults.height);
    item.aspectRatio = Number(item.aspectRatio) || width / height;
  }
  return item;
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

function escapeSvgText(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;");
}

function qrSvgDataUri(text) {
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

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${size} ${size}" shape-rendering="crispEdges"><rect width="${size}" height="${size}" fill="#ffffff"/>${cells.join("")}</svg>`;
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
}

function barcodeSvgDataUri(text) {
  const content = text || "BARCODE";
  let x = 4;
  const bars = [];
  for (const char of content) {
    const code = char.charCodeAt(0);
    const barWidth = (code % 3) + 1;
    const gap = ((code >> 2) % 2) + 1;
    bars.push(`<rect x="${x}" y="4" width="${barWidth}" height="52" fill="#111111" />`);
    x += barWidth + gap;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${Math.max(80, x + 4)} 68"><rect width="100%" height="100%" fill="#ffffff"/>${bars.join("")}<text x="50%" y="64" text-anchor="middle" font-family="Arial" font-size="8" fill="#111111">${escapeSvgText(content)}</text></svg>`;
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
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
  canvasEl.classList.add("capture-render");
  await new Promise((resolve) => window.requestAnimationFrame(resolve));
  await new Promise((resolve) => window.requestAnimationFrame(resolve));

  try {
    const canvas = await window.html2canvas(canvasEl, {
      backgroundColor: "#ffffff",
      scale: 2,
      useCORS: true,
      ignoreElements: (element) =>
        Boolean(
          element?.classList?.contains("canvas-safe") ||
          element?.classList?.contains("guide-line") ||
          element?.classList?.contains("resize-handle") ||
          element?.classList?.contains("rotate-handle") ||
          element?.classList?.contains("rotate-line") ||
          element?.classList?.contains("delete-btn")
        ),
    });

    return canvas.toDataURL("image/png");
  } finally {
    canvasEl.classList.remove("capture-render");
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
  const width = Number(state.settings?.label_width_mm) || 40;
  const height = Number(state.settings?.label_height_mm) || 14;
  elements.designCanvas.style.setProperty("--label-aspect", `${width} / ${height}`);
  elements.designCanvas.style.aspectRatio = `${width} / ${height}`;
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
  const rect = canvas.getBoundingClientRect();
  const canvasWidth = rect.width || canvas.clientWidth || 0;
  const canvasHeight = rect.height || canvas.clientHeight || 0;
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
    el.style.padding = "4px";
    el.style.pointerEvents = "auto";
    el.style.transform = `rotate(${item.rotation || 0}deg)`;
    el.style.transformOrigin = "center center";

    if (item.type === "text") {
      el.innerText = canvasItemText(item);
      el.style.fontFamily = normalizeCanvasFontFamily(item.font || state.settings.font_name);
      el.style.textRendering = "geometricPrecision";
      const fontSize = Math.max(6, Number(item.fontSize) || 50);
      if (Number.isFinite(fontSize)) {
        el.style.fontSize = `${fontSize}px`;
      } else {
        fitTextToBox(el, item);
      }
    } else {
      const image = document.createElement("img");
      image.src = generatePreviewImage(item);
      image.alt = item.type;
      image.draggable = false;
      image.style.width = "100%";
      image.style.height = "100%";
      image.style.objectFit = "contain";
      el.appendChild(image);
    }

    if (index === state.selectedItem) {
      el.classList.add("selected");
      const handle = document.createElement("div");
      handle.className = "resize-handle br";
      makeResizable(el, item, handle, index);
      el.appendChild(handle);

      const rotateLine = document.createElement("div");
      rotateLine.className = "rotate-line";

      const rotateHandle = document.createElement("div");
      rotateHandle.className = "rotate-handle";
      makeRotatable(el, item, rotateHandle, index);

      el.appendChild(rotateLine);
      el.appendChild(rotateHandle);

      const del = document.createElement("div");
      del.className = "delete-btn";
      del.innerHTML = "&times;";
      del.addEventListener("mousedown", (e) => {
        e.stopPropagation();
        e.preventDefault();
        deleteItem(index);
      });
      el.appendChild(del);
    }

    el.addEventListener("click", (e) => {
      e.stopPropagation();
      selectItem(index);
    });

      makeDraggable(el, item, index);
      canvas.appendChild(el);
    });

  syncSelectedItemControls();
}

function deleteItem(index) {
  if (index < 0 || index >= state.canvasItems.length) {
    return;
  }

  state.canvasItems.splice(index, 1);
  state.selectedItem = null;
  syncCanvasIntoSettings();
  renderCanvasFromState();
  queueSave();
  triggerPreviewAfterInteraction();
}

function saveDesign() {
  localStorage.setItem("design", JSON.stringify(state.canvasItems));
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
  queueSave();
  queuePreview();
}

function alignSelected(type) {
  if (state.selectedItem === null) {
    return;
  }

  state.isCanvasInteracting = true;
  const item = state.canvasItems[state.selectedItem];
  const rect = elements.designCanvas.getBoundingClientRect();
  const canvasWidth = rect.width || elements.designCanvas.clientWidth || 0;
  const canvasHeight = rect.height || elements.designCanvas.clientHeight || 0;

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

function makeDraggable(el, item, index) {
  el.addEventListener("mousedown", (e) => {
    if (e.target.closest(".resize-handle")) {
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
      offsetX: e.offsetX,
      offsetY: e.offsetY,
    };
    syncSelectedItemControls();
  });
}

function makeResizable(el, item, handle, index) {
  handle.addEventListener("mousedown", (e) => {
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

function applySnapping(item, canvasRect, el) {
  const centerX = canvasRect.width / 2;
  const centerY = canvasRect.height / 2;

  const elWidth = el.offsetWidth;
  const elHeight = el.offsetHeight;

  const itemCenterX = item.x + elWidth / 2;
  const itemCenterY = item.y + elHeight / 2;

  const snapThreshold = 10;

  if (Math.abs(itemCenterX - centerX) < snapThreshold) {
    item.x = centerX - elWidth / 2;
    elements.vGuide.classList.remove("hidden");
  } else {
    elements.vGuide.classList.add("hidden");
  }

  if (Math.abs(itemCenterY - centerY) < snapThreshold) {
    item.y = centerY - elHeight / 2;
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
    if (state.dragContext.mode === "resize") {
      let nextWidth = Math.max(40, e.clientX - rect.left - state.dragContext.item.x);
      let nextHeight = Math.max(20, e.clientY - rect.top - state.dragContext.item.y);
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
      const nextX = e.clientX - rect.left - state.dragContext.offsetX;
      const nextY = e.clientY - rect.top - state.dragContext.offsetY;

      state.dragContext.item.x = nextX;
      state.dragContext.item.y = nextY;
      applyStickyBounds(state.dragContext.item, elements.designCanvas, state.dragContext.el);
      applySnapping(state.dragContext.item, rect, state.dragContext.el);
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
  elements.printProgress.textContent = `Printing 1 / ${total}`;
}

function updatePrintProgress(current, total) {
  elements.printProgress.textContent = `Printing ${current} / ${total}`;
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
  updateCanvasSurface();
  syncForm();
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
      addTextButton: $("addTextButton"),
      addQrButton: $("addQrButton"),
      addBarcodeButton: $("addBarcodeButton"),
      addImageButton: $("addImageButton"),
      imageUpload: $("imageUpload"),
      saveDesignButton: $("saveDesignButton"),
      loadDesignButton: $("loadDesignButton"),
      customWidth: $("customWidth"),
      customHeight: $("customHeight"),
      applyCustomSize: $("applyCustomSize"),
      copies: $("copies"),
    copiesPlus: $("copiesPlus"),
    copiesMinus: $("copiesMinus"),
    portSelect: $("portSelect"),
    refreshPortsButton: $("refreshPortsButton"),
    bleDeviceSelect: $("bleDeviceSelect"),
    scanBleButton: $("scanBleButton"),
    connectBleButton: $("connectBleButton"),
    disconnectBleButton: $("disconnectBleButton"),
    testPrintButton: $("testPrintButton"),
    printButton: $("printButton"),
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
}

function updateConnectionUI() {
  const isBLE = state.settings.output_mode === "BLE";

  elements.connectionType.textContent = isBLE ? "BLE" : "COM";

  let statusText = "";
  let statusClass = "";

  if (isBLE) {
    if (state.connecting) {
      statusText = "Connecting...";
      statusClass = "status-connecting";
    } else if (state.bleState.connected) {
      statusText = `Connected (${state.bleState.name || state.bleState.address})`;
      statusClass = "status-connected";
    } else {
      statusText = "Disconnected";
      statusClass = "status-disconnected";
    }
  } else {
    if (!state.settings.port) {
      statusText = "Disconnected";
      statusClass = "status-disconnected";
    } else {
      statusText = `Ready (${state.settings.port})`;
      statusClass = "status-connected";
    }
  }

  elements.connectionStatus.textContent = statusText;
  elements.connectionStatus.className = `status-indicator ${statusClass}`;
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
  syncPorts();
  syncBleDevices();
  refreshBleStatus();
  updateConnectionUI();
  updateCanvasSurface();
  updatePrintPreviewOrientation();
  buildCanvasItemsFromSettings();
  renderCanvasFromState();
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
  state.settings.port = elements.portSelect.value;
  state.settings.baud_rate = 115200;
  state.settings.ble_device_address = elements.bleDeviceSelect.value;
  state.settings.ble_device_name =
      state.bleDevices.find((item) => item.address === elements.bleDeviceSelect.value)?.name || "";
  state.settings.ble_write_char_uuid = state.bleWritable[0]?.uuid || "";
  state.settings.output_mode = state.bleState.connected ? "BLE" : "Printer";
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
  try {
    setStatus("Updating...");
    elements.designCanvas.style.opacity = "0.6";
    elements.printPreview.style.opacity = "0.6";
    readFormIntoState();
    updateCanvasSurface();
    updatePrintPreviewOrientation();
    buildCanvasItemsFromSettings();
    renderCanvasFromState();
    const imageData = await captureCanvasImage();
    const result = await safeRequest("preview", {
      settings: buildRasterRenderSettings(),
      imagePath: imageData,
    });
    elements.printPreview.src = result.printImage;
    elements.designCanvas.style.opacity = "1";
    elements.printPreview.style.opacity = "1";
    setStatus("Preview updated");
  } catch (error) {
    elements.designCanvas.style.opacity = "1";
    elements.printPreview.style.opacity = "1";
    log(`Preview failed: ${error.message}`);
    setStatus("Preview failed");
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
  }
  if (elements.portSelect.value) {
    state.settings.port = elements.portSelect.value;
    queueSave();
  }
  if (state.settings.port) {
    log(`Using COM port: ${state.settings.port}`);
  }
  if (state.settings.output_mode !== "BLE" && state.settings.port) {
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
    setStatus("Test print complete");
  } catch (error) {
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

    setBusy(true, "Printing...");
    startPrintProgress(total);
    renderCanvasFromState();
    const imageData = await captureCanvasImage();
    const rasterSettings = buildRasterRenderSettings();

    let lastResult = null;
    for (let i = 1; i <= total; i += 1) {
      lastResult = await safeRequest("print", {
        settings: { ...rasterSettings, copies: 1 },
        imagePath: imageData,
      });
      updatePrintProgress(i, total);
    }

    await refreshBleState();
    if (lastResult?.mode === "BLE") {
      log(
        `BLE print sent: protocol=${lastResult.result.protocol}, bytes=${lastResult.result.bytes_sent}, packets=${lastResult.result.packet_count}, notifications=${lastResult.result.notification_count}`
      );
    } else if (lastResult?.mode === "Printer") {
      log(`Serial print sent to ${lastResult.port}`);
    }

    stopPrintProgress();
    saveRecentLabel(state.canvasItems.find((item) => item.type === "text")?.text || "");
    addPrintHistory(true, "Print successful");
    setStatus("Printed successfully");
    showPrintSuccess();
  } catch (error) {
    log(`Print failed: ${error.message}`);
    stopPrintProgress();
    addPrintHistory(false, "Print failed");
    setStatus("Print failed");
    showError(error.message);
  } finally {
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
    state.connecting = false;
    setBusy(false);
    syncForm();
    updateConnectionUI();
    startConnectionTimer();
    queueSave();
    queuePreview();
    setStatus(`BLE connected to ${state.bleState.name || state.bleState.address}`);
  } catch (error) {
    state.connecting = false;
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
      queueSave();
      queuePreview();
    });

    elements.fontSizeInput.addEventListener("input", () => {
      if (state.selectedItem === null) {
        return;
      }
      const size = Math.max(6, Number(elements.fontSizeInput.value) || 0);
      state.canvasItems[state.selectedItem].fontSize = size;
      clearError();
      syncCanvasIntoSettings();
      renderCanvasFromState();
      queueSave();
      queuePreview();
    });

    elements.systemFontToggle.addEventListener("change", async () => {
      state.useSystemFonts = elements.systemFontToggle.checked;
      localStorage.setItem("use_system_fonts", state.useSystemFonts ? "true" : "false");
      await loadFonts();
      queueSave();
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
  elements.saveDesignButton.addEventListener("click", () => {
    readFormIntoState();
    saveDesign();
    setStatus("Design saved");
  });
  elements.loadDesignButton.addEventListener("click", () => {
    if (loadDesign()) {
      queuePreview();
      queueSave();
      setStatus("Design loaded");
    }
  });
  elements.copiesPlus.addEventListener("click", () => adjustCopies(1));
  elements.copiesMinus.addEventListener("click", () => adjustCopies(-1));
  elements.applyCustomSize.addEventListener("click", applyCustomSize);
  elements.refreshPortsButton.addEventListener("click", refreshPorts);
  elements.scanBleButton.addEventListener("click", runBleScan);
  elements.connectBleButton.addEventListener("click", runBleConnect);
  elements.disconnectBleButton.addEventListener("click", runBleDisconnect);
  elements.testPrintButton.addEventListener("click", runTestPrint);
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
    queueSave();
    queuePreview();
  });
}

async function boot() {
    loadTheme();
    initElements();
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
    state.settings.ble_device_address = "";
    state.settings.ble_device_name = "";
    state.bleState = { connected: false, address: "", name: "" };
    syncStaticChoices(initData);
    forceSimpleDefaults();
    await loadFonts();
    renderRecentLabels();
    renderPrintHistory();
    await refreshPorts();
    await refreshBleState();
    stopConnectionTimer();
    state.canvasItems = [
      makeCanvasItem("text", {
        id: 1,
        key: "layer1",
        text: "Label",
        x: 80,
        y: 80,
        width: 200,
        height: 60,
      }),
    ];
    state.selectedItem = 0;
    syncCanvasIntoSettings();
    updateCanvasSurface();
    updatePrintPreviewOrientation();
    syncForm();
    renderCanvasFromState();
    await updatePreview();
    setStatus("Ready");
  }

boot().catch((error) => {
  console.error(error);
  showError(error.message);
});
