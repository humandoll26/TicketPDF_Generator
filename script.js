const BLEED_MM = 3;
const PRESETS = {
  "55x150": { width: 55, height: 150 },
  "76x150": { width: 76, height: 150 },
  "60x170": { width: 60, height: 170 },
};

const LIMITS = {
  width: { min: 20, max: 200 },
  height: { min: 20, max: 100 },
  perforationMin: 1,
  serialCount: { min: 1, max: 1000 },
};

const SAMPLE_STATE = {
  sizePreset: "55x150",
  swapOrientation: false,
  customWidth: 55,
  customHeight: 100,
  backgroundStyle: "none",
  backgroundDirection: "vertical",
  backgroundColorA: "#ffffff",
  backgroundColorB: "#dfe6ea",
  fontColor: "#111111",
  labelColor: "#666666",
  perforationEnabled: true,
  perforationPosition: 40,
  eventName: "春の演劇祭",
  venue: "〇〇ホール",
  date: "2026.04.12",
  time: "18:30",
  seat: "自由席",
  price: "3,000円",
  notes: "開場は30分前",
  serialEnabled: true,
  serialMode: "placeholder",
  serialStart: "###",
  serialCount: 10,
  fontSizeEventName: 24,
  fontSizeDate: 10,
  fontSizeTime: 10,
  fontSizeVenue: 10,
  fontSizeSeat: 16,
  fontSizePrice: 10,
  fontSizeNotes: 8,
  fontSizeStubEventName: 16,
  fontSizeStubDate: 8,
  fontSizeStubTime: 8,
  fontSizeSerial: 16,
};

const ui = {
  form: document.getElementById("ticket-form"),
  sizePreset: document.getElementById("sizePreset"),
  customSizeFields: document.getElementById("customSizeFields"),
  swapOrientation: document.getElementById("swapOrientation"),
  customWidth: document.getElementById("customWidth"),
  customHeight: document.getElementById("customHeight"),
  backgroundStyle: document.getElementById("backgroundStyle"),
  backgroundDirection: document.getElementById("backgroundDirection"),
  backgroundColorA: document.getElementById("backgroundColorA"),
  backgroundColorB: document.getElementById("backgroundColorB"),
  fontColor: document.getElementById("fontColor"),
  labelColor: document.getElementById("labelColor"),
  perforationEnabled: document.getElementById("perforationEnabled"),
  perforationFields: document.getElementById("perforationFields"),
  perforationLabel: document.getElementById("perforationLabel"),
  perforationPrintNote: document.getElementById("perforationPrintNote"),
  perforationPrintComment: document.getElementById("perforationPrintComment"),
  perforationPosition: document.getElementById("perforationPosition"),
  serialEnabled: document.getElementById("serialEnabled"),
  serialFields: document.getElementById("serialFields"),
  serialMode: document.getElementById("serialMode"),
  serialStartLabel: document.getElementById("serialStartLabel"),
  serialModeHint: document.getElementById("serialModeHint"),
  serialStart: document.getElementById("serialStart"),
  serialCount: document.getElementById("serialCount"),
  serialPrintComment: document.getElementById("serialPrintComment"),
  fontSizeEventName: document.getElementById("fontSizeEventName"),
  fontSizeDate: document.getElementById("fontSizeDate"),
  fontSizeTime: document.getElementById("fontSizeTime"),
  fontSizeVenue: document.getElementById("fontSizeVenue"),
  fontSizeSeat: document.getElementById("fontSizeSeat"),
  fontSizePrice: document.getElementById("fontSizePrice"),
  fontSizeNotes: document.getElementById("fontSizeNotes"),
  fontSizeStubEventName: document.getElementById("fontSizeStubEventName"),
  fontSizeStubDate: document.getElementById("fontSizeStubDate"),
  fontSizeStubTime: document.getElementById("fontSizeStubTime"),
  fontSizeSerial: document.getElementById("fontSizeSerial"),
  eventName: document.getElementById("eventName"),
  venue: document.getElementById("venue"),
  date: document.getElementById("date"),
  time: document.getElementById("time"),
  seat: document.getElementById("seat"),
  price: document.getElementById("price"),
  notes: document.getElementById("notes"),
  errorBox: document.getElementById("errorBox"),
  canvas: document.getElementById("previewCanvas"),
  preflightChecklist: document.getElementById("preflightChecklist"),
  checklistStatus: document.getElementById("checklistStatus"),
  generateButton: document.getElementById("generatePdfButton"),
  exportSettingsButton: document.getElementById("exportSettingsButton"),
  importSettingsButton: document.getElementById("importSettingsButton"),
  importSettingsInput: document.getElementById("importSettingsInput"),
};

const ctx = ui.canvas.getContext("2d");
const perforationPositions = {
  horizontal: null,
  vertical: null,
};
let lastOrientation = null;
const accordions = Array.from(document.querySelectorAll("[data-accordion]"));
const STORAGE_KEY = "ticket-pdf-generator-state-v2";
const STORAGE_SCHEMA_VERSION = 2;
const LEGACY_STORAGE_KEYS = ["ticket-pdf-generator-state-v1"];
const FONT_SIZE_KEYS = [
  "fontSizeEventName",
  "fontSizeDate",
  "fontSizeTime",
  "fontSizeVenue",
  "fontSizeSeat",
  "fontSizePrice",
  "fontSizeNotes",
  "fontSizeStubEventName",
  "fontSizeStubDate",
  "fontSizeStubTime",
  "fontSizeSerial",
];
const TEXT_FIELD_KEYS = ["eventName", "venue", "date", "time", "seat", "price", "notes", "serialStart"];
const PREFLIGHT_CHECKS = [
  { id: "size", label: "仕上がりサイズを確認した" },
  { id: "bleed", label: "塗り足し3mmを確認した" },
  { id: "perforation", label: "ミシン目の有無と位置を確認した", when: (state) => state.perforationEnabled },
  { id: "serial", label: "連番の運用方法を確認した", when: (state) => state.serialEnabled },
  { id: "instructions", label: "印刷所への指示文を確認した", when: (state) => state.perforationEnabled || state.serialEnabled },
  { id: "preview", label: "最終プレビューを確認した" },
];
let preflightChecklistState = {};

init();

function init() {
  applyStateToForm(SAMPLE_STATE);
  restoreSavedState();
  bindEvents();
  updateColorSwatches();
  updateFormVisibility();
  renderAll();
}

function applyStateToForm(state) {
  Object.entries(state).forEach(([key, value]) => {
    const element = ui[key];
    if (!element) {
      return;
    }

    if (element.type === "checkbox") {
      element.checked = Boolean(value);
    } else {
      element.value = value;
    }
  });
}

function bindEvents() {
  accordions.forEach((accordion) => {
    const toggle = accordion.querySelector("[data-accordion-toggle]");
    const panel = accordion.querySelector("[data-accordion-panel]");
    if (!toggle || !panel) {
      return;
    }

    toggle.addEventListener("click", () => {
      const willOpen = !accordion.classList.contains("is-open");
      accordion.classList.toggle("is-open", willOpen);
      toggle.setAttribute("aria-expanded", String(willOpen));
      panel.hidden = !willOpen;
    });
  });

  ui.form.addEventListener("input", () => {
    updateColorSwatches();
    updateFormVisibility();
    syncPreflightChecklist();
    persistState();
    renderAll();
  });

  ui.form.addEventListener("change", () => {
    updateColorSwatches();
    updateFormVisibility();
    syncPreflightChecklist();
    persistState();
    renderAll();
  });

  ui.generateButton.addEventListener("click", generatePdf);
  window.addEventListener("resize", renderPreview);
  document.getElementById("resetFontSizesButton").addEventListener("click", resetFontSizes);
  document.getElementById("clearSavedDataButton").addEventListener("click", clearSavedData);
  ui.exportSettingsButton.addEventListener("click", exportSettings);
  ui.importSettingsButton.addEventListener("click", () => ui.importSettingsInput.click());
  ui.importSettingsInput.addEventListener("change", importSettings);
  ui.serialMode.addEventListener("change", () => {
    if (ui.serialMode.value === "placeholder" && !/^#+$/.test(ui.serialStart.value.trim())) {
      ui.serialStart.value = "###";
    }
    if (ui.serialMode.value === "numbered" && /^#+$/.test(ui.serialStart.value.trim())) {
      ui.serialStart.value = "001";
    }
  });

  ui.perforationPosition.addEventListener("input", () => {
    const ticket = getTicketDimensions(getFormState());
    if (!ticket) {
      return;
    }
    perforationPositions[getTicketOrientation(ticket)] = Number(ui.perforationPosition.value);
  });

  ui.preflightChecklist.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "checkbox") {
      return;
    }

    preflightChecklistState[target.value] = target.checked;
    updateGenerateButtonState();
  });
}

function updateFormVisibility() {
  const state = getFormState();
  const ticket = getTicketDimensions(state);
  const useCustom = state.sizePreset === "custom";
  ui.customSizeFields.classList.toggle("is-hidden", !useCustom);
  ui.customSizeFields.setAttribute("aria-hidden", String(!useCustom));

  const showPerforation = state.perforationEnabled;
  ui.perforationFields.classList.toggle("is-hidden", !showPerforation);
  ui.perforationPrintNote.classList.toggle("is-hidden", !showPerforation);
  ui.perforationPosition.disabled = !showPerforation;
  ui.perforationLabel.textContent = getPerforationLabel(ticket);
  if (showPerforation) {
    ui.perforationPrintComment.textContent = getPerforationPrintComment(ticket);
  }
  syncPerforationPosition(ticket);

  const showSerial = state.serialEnabled;
  ui.serialFields.classList.toggle("is-hidden", !showSerial);
  ui.serialMode.disabled = !showSerial;
  ui.serialStart.disabled = !showSerial;
  ui.serialCount.disabled = !showSerial;
  const serialConfig = resolveSerialSettings(state);
  ui.serialStartLabel.textContent = serialConfig.mode === "placeholder" ? "差し込み記号" : "開始番号";
  ui.serialModeHint.textContent = getSerialModeHint(serialConfig.mode);
  ui.serialStart.placeholder = serialConfig.mode === "placeholder" ? "###" : "001";
  ui.serialPrintComment.textContent = getSerialPrintComment(state);
  renderPreflightChecklist(state);
}

function getFormState() {
  return normalizeState(readFormValues());
}

function readFormValues() {
  return {
    sizePreset: ui.sizePreset.value,
    customWidth: Number(ui.customWidth.value),
    customHeight: Number(ui.customHeight.value),
    swapOrientation: ui.swapOrientation.checked,
    perforationEnabled: ui.perforationEnabled.checked,
    backgroundStyle: ui.backgroundStyle.value,
    backgroundDirection: ui.backgroundDirection.value,
    backgroundColorA: ui.backgroundColorA.value,
    backgroundColorB: ui.backgroundColorB.value,
    fontColor: ui.fontColor.value,
    labelColor: ui.labelColor.value,
    perforationPosition: Number(ui.perforationPosition.value),
    eventName: ui.eventName.value.trim(),
    venue: ui.venue.value.trim(),
    date: ui.date.value.trim(),
    time: ui.time.value.trim(),
    seat: ui.seat.value.trim(),
    price: ui.price.value.trim(),
    notes: ui.notes.value.trim(),
    serialEnabled: ui.serialEnabled.checked,
    serialMode: ui.serialMode.value,
    serialStart: ui.serialStart.value.trim(),
    serialCount: Number(ui.serialCount.value),
    fontSizeEventName: Number(ui.fontSizeEventName.value),
    fontSizeDate: Number(ui.fontSizeDate.value),
    fontSizeTime: Number(ui.fontSizeTime.value),
    fontSizeVenue: Number(ui.fontSizeVenue.value),
    fontSizeSeat: Number(ui.fontSizeSeat.value),
    fontSizePrice: Number(ui.fontSizePrice.value),
    fontSizeNotes: Number(ui.fontSizeNotes.value),
    fontSizeStubEventName: Number(ui.fontSizeStubEventName.value),
    fontSizeStubDate: Number(ui.fontSizeStubDate.value),
    fontSizeStubTime: Number(ui.fontSizeStubTime.value),
    fontSizeSerial: Number(ui.fontSizeSerial.value),
  };
}

function normalizeState(rawState = {}) {
  const merged = {
    ...SAMPLE_STATE,
    ...rawState,
  };

  const normalized = {
    ...merged,
    customWidth: toNumber(merged.customWidth, SAMPLE_STATE.customWidth),
    customHeight: toNumber(merged.customHeight, SAMPLE_STATE.customHeight),
    perforationPosition: toNumber(merged.perforationPosition, SAMPLE_STATE.perforationPosition),
    serialCount: toInteger(merged.serialCount, SAMPLE_STATE.serialCount),
    swapOrientation: Boolean(merged.swapOrientation),
    perforationEnabled: Boolean(merged.perforationEnabled),
    serialEnabled: Boolean(merged.serialEnabled),
    backgroundStyle: normalizeEnum(merged.backgroundStyle, ["none", "solid", "gradient"], SAMPLE_STATE.backgroundStyle),
    backgroundDirection: normalizeEnum(merged.backgroundDirection, ["vertical", "horizontal", "diagonal"], SAMPLE_STATE.backgroundDirection),
    sizePreset: normalizeEnum(merged.sizePreset, [...Object.keys(PRESETS), "custom"], SAMPLE_STATE.sizePreset),
  };

  TEXT_FIELD_KEYS.forEach((key) => {
    normalized[key] = typeof merged[key] === "string" ? merged[key].trim() : String(merged[key] ?? SAMPLE_STATE[key]).trim();
  });

  FONT_SIZE_KEYS.forEach((key) => {
    normalized[key] = toInteger(merged[key], SAMPLE_STATE[key]);
  });

  normalized.serialMode = normalizeEnum(merged.serialMode, ["placeholder", "numbered"], normalizeSerialMode(normalized));

  if (normalized.serialMode === "placeholder" && !/^#+$/.test(normalized.serialStart || "")) {
    normalized.serialStart = "###";
  }

  if (normalized.serialMode === "numbered" && /^#+$/.test(normalized.serialStart || "")) {
    normalized.serialStart = "001";
  }

  return normalized;
}

function buildStoragePayload(state) {
  return {
    schemaVersion: STORAGE_SCHEMA_VERSION,
    savedAt: new Date().toISOString(),
    state,
  };
}

function parseStoredState(payload) {
  if (!payload || typeof payload !== "object") {
    return normalizeState(SAMPLE_STATE);
  }

  if ("schemaVersion" in payload) {
    return normalizeState(migrateState(payload));
  }

  return normalizeState(migrateState({
    schemaVersion: 1,
    state: payload,
  }));
}

function migrateState(payload) {
  const schemaVersion = Number(payload.schemaVersion || 1);
  let state = { ...(payload.state || {}) };

  if (schemaVersion < 2) {
    state = {
      ...state,
      serialMode: state.serialMode || (/^#+$/.test(state.serialStart || "") ? "placeholder" : "numbered"),
    };
  }

  return state;
}

function persistState() {
  try {
    const state = getFormState();
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(buildStoragePayload(state)));
    updateSaveStatus("入力内容を保存しました");
  } catch (error) {
    console.error(error);
    updateSaveStatus("保存に失敗しました");
  }
}

function restoreSavedState() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY) || loadLegacyStoredState();
    if (!raw) {
      return;
    }

    const saved = parseStoredState(JSON.parse(raw));
    applyStateToForm(saved);

    updateSaveStatus("保存済みの入力内容を読み込みました");
    updateColorSwatches();
  } catch (error) {
    console.error(error);
    updateSaveStatus("保存データを読み込めませんでした");
  }
}

function resetFontSizes() {
  FONT_SIZE_KEYS.forEach((key) => {
    const element = ui[key];
    if (!element) {
      return;
    }
    element.value = String(SAMPLE_STATE[key]);
  });

  persistState();
  renderAll();
  updateColorSwatches();
  updateSaveStatus("フォントサイズを初期値に戻しました");
}

function updateSaveStatus(message) {
  const status = document.getElementById("saveStatus");
  if (!status) {
    return;
  }
  status.textContent = message;
}

function updateColorSwatches() {
  const entries = [
    ["backgroundColorASwatch", ui.backgroundColorA?.value],
    ["backgroundColorBSwatch", ui.backgroundColorB?.value],
    ["fontColorSwatch", ui.fontColor?.value],
    ["labelColorSwatch", ui.labelColor?.value],
  ];

  entries.forEach(([id, color]) => {
    const swatch = document.getElementById(id);
    if (!swatch || !color) {
      return;
    }
    swatch.style.background = color;
  });
}

function clearSavedData() {
  const confirmed = window.confirm("保存済みの入力内容を初期化します。よろしいですか？");
  if (!confirmed) {
    return;
  }

  try {
    window.localStorage.removeItem(STORAGE_KEY);
    LEGACY_STORAGE_KEYS.forEach((key) => window.localStorage.removeItem(key));
    applyStateToForm(SAMPLE_STATE);
    preflightChecklistState = {};

    perforationPositions.horizontal = null;
    perforationPositions.vertical = null;
    lastOrientation = null;
    updateColorSwatches();
    updateFormVisibility();
    renderAll();
    updateSaveStatus("保存内容を初期化しました");
  } catch (error) {
    console.error(error);
    updateSaveStatus("保存内容の初期化に失敗しました");
  }
}

function loadLegacyStoredState() {
  for (const key of LEGACY_STORAGE_KEYS) {
    const raw = window.localStorage.getItem(key);
    if (raw) {
      return raw;
    }
  }
  return null;
}

function exportSettings() {
  try {
    const state = getFormState();
    const blob = new Blob([JSON.stringify(buildStoragePayload(state), null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "ticket-settings.json";
    link.click();
    URL.revokeObjectURL(url);
    updateSaveStatus("設定ファイルを書き出しました");
  } catch (error) {
    console.error(error);
    updateSaveStatus("設定ファイルの書き出しに失敗しました");
  }
}

async function importSettings(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    const payload = JSON.parse(await file.text());
    const state = parseStoredState(payload);
    applyStateToForm(state);
    updateColorSwatches();
    updateFormVisibility();
    syncPreflightChecklist();
    renderAll();
    persistState();
    updateSaveStatus("設定ファイルを読み込みました");
  } catch (error) {
    console.error(error);
    renderErrors(["設定ファイルを読み込めませんでした。JSON形式を確認してください。"]);
    updateSaveStatus("設定ファイルの読み込みに失敗しました");
  } finally {
    ui.importSettingsInput.value = "";
  }
}

function renderAll() {
  const state = getFormState();
  const errors = validateState(state);
  renderErrors(errors);
  updateGenerateButtonState(errors.length === 0);
  renderPreview();
}

function validateState(state) {
  const errors = [];
  const ticket = getTicketDimensions(state);

  if (state.sizePreset === "custom") {
    if (!isBetween(state.customWidth, LIMITS.width.min, LIMITS.width.max)) {
      errors.push(`カスタム幅は ${LIMITS.width.min}〜${LIMITS.width.max}mm で入力してください。`);
    }
    if (!isBetween(state.customHeight, LIMITS.height.min, LIMITS.height.max)) {
      errors.push(`カスタム高さは ${LIMITS.height.min}〜${LIMITS.height.max}mm で入力してください。`);
    }
  }

  if (!ticket) {
    return [...new Set(errors)];
  }

  if (state.perforationEnabled) {
    const maxPosition = getPerforationLimit(ticket);
    if (!(state.perforationPosition >= LIMITS.perforationMin && state.perforationPosition < maxPosition)) {
      errors.push(`ミシン位置は 1mm 以上かつ ${maxPosition}mm 未満で入力してください。`);
    }
  }

  if (state.serialEnabled) {
    const serialConfig = resolveSerialSettings(state);
    if (!serialConfig.isValid) {
      errors.push(serialConfig.errorMessage);
    }
    if (!Number.isInteger(state.serialCount) || !isBetween(state.serialCount, LIMITS.serialCount.min, LIMITS.serialCount.max)) {
      errors.push(`発行枚数は ${LIMITS.serialCount.min}〜${LIMITS.serialCount.max} の整数で入力してください。`);
    }
  }

  return [...new Set(errors)];
}

function renderErrors(errors) {
  if (!errors.length) {
    ui.errorBox.classList.remove("is-visible");
    ui.errorBox.textContent = "";
    return;
  }

  ui.errorBox.classList.add("is-visible");
  ui.errorBox.innerHTML = errors.map((error) => `<div>${escapeHtml(error)}</div>`).join("");
}

function getVisiblePreflightChecks(state) {
  return PREFLIGHT_CHECKS.filter((item) => (typeof item.when === "function" ? item.when(state) : true));
}

function renderPreflightChecklist(state) {
  const visibleChecks = getVisiblePreflightChecks(state);
  const nextState = {};

  ui.preflightChecklist.innerHTML = visibleChecks.map((item) => {
    nextState[item.id] = Boolean(preflightChecklistState[item.id]);
    return `
      <label class="checklist-item">
        <input type="checkbox" value="${escapeHtml(item.id)}" ${nextState[item.id] ? "checked" : ""}>
        <span>${escapeHtml(item.label)}</span>
      </label>
    `;
  }).join("");

  preflightChecklistState = nextState;
  updateGenerateButtonState();
}

function syncPreflightChecklist() {
  const state = getFormState();
  const visibleIds = new Set(getVisiblePreflightChecks(state).map((item) => item.id));
  Object.keys(preflightChecklistState).forEach((id) => {
    if (!visibleIds.has(id)) {
      delete preflightChecklistState[id];
    }
  });
  renderPreflightChecklist(state);
}

function updateGenerateButtonState(isFormValid = !ui.errorBox.classList.contains("is-visible")) {
  const requiredChecks = Object.keys(preflightChecklistState);
  const allChecked = requiredChecks.every((id) => preflightChecklistState[id]);
  ui.generateButton.disabled = !(isFormValid && allChecked);
  ui.checklistStatus.textContent = allChecked
    ? "確認済みです。PDFを生成できます。"
    : "必要な項目を確認するとPDF生成できます。";
}

function renderPreview() {
  const state = getFormState();
  const ticket = getTicketDimensions(state);

  const ratio = window.devicePixelRatio || 1;
  const wrapWidth = ui.canvas.parentElement ? ui.canvas.parentElement.clientWidth : 960;
  const cssWidth = Math.max(320, Math.floor(wrapWidth || 960));
  const cssHeight = Math.max(420, Math.floor(cssWidth * 0.66));

  ui.canvas.width = cssWidth * ratio;
  ui.canvas.height = cssHeight * ratio;
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0);

  ctx.clearRect(0, 0, cssWidth, cssHeight);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, cssWidth, cssHeight);
  ui.canvas.style.width = `${cssWidth}px`;
  ui.canvas.style.height = `${cssHeight}px`;

  if (!ticket) {
    drawPreviewPlaceholder(cssWidth, cssHeight);
    return;
  }

  const bleedBox = {
    width: ticket.width + BLEED_MM * 2,
    height: ticket.height + BLEED_MM * 2,
  };

  const margin = 48;
  const scale = Math.min((cssWidth - margin * 2) / bleedBox.width, (cssHeight - margin * 2) / bleedBox.height);
  const drawWidth = bleedBox.width * scale;
  const drawHeight = bleedBox.height * scale;
  const originX = (cssWidth - drawWidth) / 2;
  const originY = (cssHeight - drawHeight) / 2;

  const geometry = {
    cssWidth,
    cssHeight,
    scale,
    originX,
    originY,
    bleedPx: BLEED_MM * scale,
    trimX: originX + BLEED_MM * scale,
    trimY: originY + BLEED_MM * scale,
    trimWidth: ticket.width * scale,
    trimHeight: ticket.height * scale,
    bleedWidth: drawWidth,
    bleedHeight: drawHeight,
    ticket,
  };

  drawBleedArea(geometry);
  fillTicketTexture(geometry, state);
  drawTrimLine(geometry);
  drawCropMarks(geometry);
  drawPerforation(geometry, state);
  drawTicketTextCanvas(geometry, state, getSerialForIndex(state, 0));
}

function drawPreviewPlaceholder(width, height, targetCtx = ctx) {
  targetCtx.strokeStyle = "#d0d0d0";
  targetCtx.strokeRect(24, 24, width - 48, height - 48);
  targetCtx.fillStyle = "#666";
  targetCtx.font = "16px 'IBM Plex Sans JP', sans-serif";
  targetCtx.textAlign = "center";
  targetCtx.fillText("サイズ設定を見直してください", width / 2, height / 2);
}

function drawBleedArea(geometry, targetCtx = ctx) {
  const { originX, originY, bleedWidth, bleedHeight } = geometry;
  targetCtx.fillStyle = "#ffffff";
  targetCtx.fillRect(originX, originY, bleedWidth, bleedHeight);
  targetCtx.strokeStyle = "#d97373";
  targetCtx.lineWidth = 1.4;
  targetCtx.strokeRect(originX, originY, bleedWidth, bleedHeight);
}

function drawTrimLine(geometry, targetCtx = ctx) {
  const { trimX, trimY, trimWidth, trimHeight } = geometry;
  targetCtx.strokeStyle = "#111111";
  targetCtx.lineWidth = 1.2;
  targetCtx.strokeRect(trimX, trimY, trimWidth, trimHeight);
}

function fillTicketTexture(geometry, state, targetCtx = ctx) {
  const { trimX, trimY, trimWidth, trimHeight } = geometry;
  const backgroundStyle = state.backgroundStyle || "none";
  const colorA = state.backgroundColorA || "#ffffff";
  const colorB = state.backgroundColorB || "#dfe6ea";

  targetCtx.save();
  targetCtx.fillStyle = "#ffffff";
  targetCtx.fillRect(trimX, trimY, trimWidth, trimHeight);

  if (backgroundStyle === "solid") {
    targetCtx.fillStyle = colorA;
    targetCtx.fillRect(trimX, trimY, trimWidth, trimHeight);
  } else if (backgroundStyle === "gradient") {
    let gradient;
    if (state.backgroundDirection === "horizontal") {
      gradient = targetCtx.createLinearGradient(trimX, trimY, trimX + trimWidth, trimY);
    } else if (state.backgroundDirection === "diagonal") {
      gradient = targetCtx.createLinearGradient(trimX, trimY, trimX + trimWidth, trimY + trimHeight);
    } else {
      gradient = targetCtx.createLinearGradient(trimX, trimY, trimX, trimY + trimHeight);
    }
    gradient.addColorStop(0, colorA);
    gradient.addColorStop(1, colorB);
    targetCtx.fillStyle = gradient;
    targetCtx.fillRect(trimX, trimY, trimWidth, trimHeight);
  }

  targetCtx.restore();
}

function drawCropMarks(geometry, targetCtx = ctx) {
  const { trimX, trimY, trimWidth, trimHeight, scale } = geometry;
  const gap = 0.8 * scale;
  const length = 1.6 * scale;

  targetCtx.strokeStyle = "#111111";
  targetCtx.lineWidth = 1;

  const corners = [
    { x: trimX, y: trimY, hx: -1, hy: -1 },
    { x: trimX + trimWidth, y: trimY, hx: 1, hy: -1 },
    { x: trimX, y: trimY + trimHeight, hx: -1, hy: 1 },
    { x: trimX + trimWidth, y: trimY + trimHeight, hx: 1, hy: 1 },
  ];

  corners.forEach((corner) => {
    targetCtx.beginPath();
    targetCtx.moveTo(corner.x + corner.hx * gap, corner.y + corner.hy * (gap + length));
    targetCtx.lineTo(corner.x + corner.hx * gap, corner.y + corner.hy * gap);
    targetCtx.lineTo(corner.x + corner.hx * (gap + length), corner.y + corner.hy * gap);
    targetCtx.stroke();
  });
}

function drawPerforation(geometry, state, targetCtx = ctx) {
  if (!state.perforationEnabled) {
    return;
  }

  const { trimX, trimY, trimHeight, trimWidth, scale, ticket } = geometry;
  const orientation = getTicketOrientation(ticket);

  targetCtx.save();
  targetCtx.strokeStyle = "rgba(17, 17, 17, 0.55)";
  targetCtx.lineWidth = 0.8;
  targetCtx.setLineDash([3, 4]);
  targetCtx.beginPath();

  if (orientation === "horizontal") {
    const x = trimX + state.perforationPosition * scale;
    targetCtx.moveTo(x, trimY);
    targetCtx.lineTo(x, trimY + trimHeight);
  } else {
    const y = trimY + state.perforationPosition * scale;
    targetCtx.moveTo(trimX, y);
    targetCtx.lineTo(trimX + trimWidth, y);
  }

  targetCtx.stroke();
  targetCtx.restore();
}

function drawTicketTextCanvas(geometry, state, serial, targetCtx = ctx) {
  const { scale, ticket } = geometry;
  const areas = getTicketAreas(geometry, state);
  const orientation = getTicketOrientation(ticket);
  const layout = buildTicketLayout(areas, state, scale, orientation, Boolean(serial));

  targetCtx.fillStyle = state.fontColor || "#111111";
  drawMainTicketLayout(targetCtx, layout.main, state, scale, orientation);

  if (layout.stub) {
    drawStubLayout(targetCtx, layout.stub, state, scale, orientation, serial);
  }
}

function drawWrappedTextBlock(context, options) {
  const {
    text,
    x,
    y,
    maxWidth,
    maxLines,
    lineHeight,
    fontSize,
    fontFamily,
    fontWeight,
    align,
  } = options;

  if (!text) {
    return 0;
  }

  context.save();
  context.textAlign = align;
  context.font = `${fontWeight} ${fontSize}px ${fontFamily}`;

  const lines = wrapTextByWidth(context, text, maxWidth).slice(0, maxLines);
  lines.forEach((line, index) => {
    context.fillText(line, x, y + lineHeight * (index + 1));
  });
  context.restore();

  return lineHeight * lines.length;
}

function wrapTextByWidth(context, text, maxWidth) {
  if (!text) {
    return [];
  }

  const lines = [];
  let current = "";

  Array.from(text).forEach((char) => {
    const test = current + char;
    if (context.measureText(test).width > maxWidth && current) {
      lines.push(current);
      current = char;
    } else {
      current = test;
    }
  });

  if (current) {
    lines.push(current);
  }

  return lines;
}

function getTicketDimensions(state) {
  let dimensions;

  if (state.sizePreset !== "custom") {
    dimensions = PRESETS[state.sizePreset];
  } else {
    if (!isBetween(state.customWidth, LIMITS.width.min, LIMITS.width.max) || !isBetween(state.customHeight, LIMITS.height.min, LIMITS.height.max)) {
      return null;
    }

    dimensions = {
      width: state.customWidth,
      height: state.customHeight,
    };
  }

  if (state.swapOrientation) {
    return {
      width: dimensions.height,
      height: dimensions.width,
    };
  }

  return dimensions;
}

function getTicketOrientation(ticket) {
  return ticket.width > ticket.height ? "horizontal" : "vertical";
}

function getPerforationLimit(ticket) {
  return getTicketOrientation(ticket) === "horizontal" ? ticket.width : ticket.height;
}

function getRecommendedPerforationPosition(ticket) {
  const orientation = getTicketOrientation(ticket);

  if (orientation === "horizontal") {
    const stubWidth = clamp(ticket.width * 0.24, 32, 42);
    return Math.round(ticket.width - stubWidth);
  }

  const stubHeight = clamp(ticket.height * 0.2, 26, 36);
  return Math.round(ticket.height - stubHeight);
}

function getPerforationLabel(ticket) {
  if (!ticket) {
    return "ミシン位置（mm）";
  }

  return getTicketOrientation(ticket) === "horizontal"
    ? "ミシン位置（左から mm）"
    : "ミシン位置（上から mm）";
}

function getPerforationPrintComment(ticket) {
  if (!ticket) {
    return "仕上がり位置に合わせてミシン目加工を1本お願いします。";
  }

  const orientation = getTicketOrientation(ticket);
  const axis = orientation === "horizontal" ? "左から" : "上から";
  const position = Number(ui.perforationPosition.value || getRecommendedPerforationPosition(ticket));
  return `仕上がり位置で ${axis}${position}mm の位置にミシン目加工を1本お願いします。`;
}

function syncPerforationPosition(ticket) {
  if (!ticket) {
    return;
  }

  const orientation = getTicketOrientation(ticket);
  const limit = getPerforationLimit(ticket);

  if (lastOrientation && lastOrientation !== orientation) {
    perforationPositions[lastOrientation] = Number(ui.perforationPosition.value);
  }

  let nextValue = perforationPositions[orientation];

  if (!Number.isFinite(nextValue)) {
    nextValue = getRecommendedPerforationPosition(ticket);
  }

  nextValue = clamp(Math.round(nextValue), LIMITS.perforationMin, limit - 1);
  perforationPositions[orientation] = nextValue;

  if (Number(ui.perforationPosition.value) !== nextValue) {
    ui.perforationPosition.value = String(nextValue);
  }

  lastOrientation = orientation;
}

function getTicketAreas(geometry, state) {
  const { trimX, trimY, trimWidth, trimHeight, scale, ticket } = geometry;
  const padding = 6 * scale;
  const orientation = getTicketOrientation(ticket);

  if (!state.perforationEnabled) {
    return {
      main: {
        x: trimX + padding,
        y: trimY + padding,
        width: trimWidth - padding * 2,
        height: trimHeight - padding * 2,
      },
      stub: null,
      orientation,
    };
  }

  if (orientation === "horizontal") {
    const splitX = trimX + state.perforationPosition * scale;
    return {
      main: {
        x: trimX + padding,
        y: trimY + padding,
        width: Math.max(24, splitX - trimX - padding * 1.5),
        height: trimHeight - padding * 2,
      },
      stub: {
        x: splitX + padding * 0.5,
        y: trimY + padding,
        width: Math.max(18, trimX + trimWidth - splitX - padding * 1.5),
        height: trimHeight - padding * 2,
      },
      orientation,
    };
  }

  const splitY = trimY + state.perforationPosition * scale;
  return {
    main: {
      x: trimX + padding,
      y: trimY + padding,
      width: trimWidth - padding * 2,
      height: Math.max(28, splitY - trimY - padding * 1.5),
    },
    stub: {
      x: trimX + padding,
      y: splitY + padding * 0.5,
      width: trimWidth - padding * 2,
      height: Math.max(18, trimY + trimHeight - splitY - padding * 1.5),
    },
    orientation,
  };
}

function buildTicketLayout(areas, state, scale, orientation, hasSerial) {
  return {
    main: orientation === "horizontal"
      ? buildHorizontalMainLayout(areas.main, state, scale)
      : buildVerticalMainLayout(areas.main, state, scale),
    stub: areas.stub
      ? (orientation === "horizontal"
          ? buildHorizontalStubLayout(areas.stub, scale, hasSerial)
          : buildVerticalStubLayout(areas.stub, scale, hasSerial))
      : null,
  };
}

function buildHorizontalMainLayout(area, state, scale) {
  const inner = insetRect(area, 5 * scale, 4 * scale);
  const leftWidth = inner.width * 0.6;
  const rightWidth = inner.width - leftWidth - 5 * scale;
  const titleSize = mmFontToPx(state.fontSizeEventName, scale, 1.05);
  const metaSize = Math.max(
    mmFontToPx(state.fontSizeDate, scale),
    mmFontToPx(state.fontSizeTime, scale),
    mmFontToPx(state.fontSizeVenue, scale)
  );
  const notesSize = mmFontToPx(state.fontSizeNotes, scale);
  const titleRect = {
    x: inner.x,
    y: inner.y,
    width: leftWidth,
    height: Math.min(inner.height * 0.46, titleSize * 2.5 + 7 * scale),
  };
  const metaRect = {
    x: inner.x + leftWidth + 5 * scale,
    y: inner.y,
    width: rightWidth,
    height: Math.min(inner.height * 0.5, metaSize * 4.2 + 10 * scale),
  };
  const notesHeight = state.notes ? Math.max(notesSize * 2.8 + 6 * scale, inner.height * 0.18) : 0;
  const detailsY = Math.max(titleRect.y + titleRect.height + 5 * scale, metaRect.y + metaRect.height + 4 * scale);

  return {
    title: titleRect,
    meta: metaRect,
    details: {
      x: inner.x,
      y: detailsY,
      width: inner.width,
      height: Math.max(8 * scale, inner.height - (detailsY - inner.y) - notesHeight - 3 * scale),
    },
    notes: state.notes
      ? {
          x: inner.x,
          y: inner.y + inner.height - notesHeight,
          width: inner.width,
          height: notesHeight,
        }
      : null,
  };
}

function buildVerticalMainLayout(area, state, scale) {
  const inner = insetRect(area, 5 * scale, 4 * scale);
  const titleHeight = clamp(inner.height * 0.24, 16 * scale, 30 * scale);
  const notesHeight = state.notes ? clamp(inner.height * 0.16, 8 * scale, 18 * scale) : 0;
  const detailsTop = inner.y + titleHeight + 6 * scale;
  return {
    title: {
      x: inner.x,
      y: inner.y,
      width: inner.width,
      height: titleHeight,
    },
    details: {
      x: inner.x,
      y: detailsTop,
      width: inner.width,
      height: Math.max(10 * scale, inner.y + inner.height - notesHeight - 2 * scale - detailsTop),
    },
    notes: state.notes
      ? {
          x: inner.x,
          y: inner.y + inner.height - notesHeight,
          width: inner.width,
          height: notesHeight,
        }
      : null,
  };
}

function buildHorizontalStubLayout(area, scale, hasSerial) {
  const inner = insetRect(area, 3.5 * scale, 3.5 * scale);
  return {
    info: inner,
    serial: hasSerial
      ? {
          x: inner.x,
          y: inner.y + inner.height - 10 * scale,
          width: inner.width,
          height: 10 * scale,
        }
      : null,
  };
}

function buildVerticalStubLayout(area, scale, hasSerial) {
  const inner = insetRect(area, 3.5 * scale, 3.5 * scale);
  const serialHeight = hasSerial ? clamp(inner.height * 0.2, 8 * scale, 14 * scale) : 0;
  return {
    info: {
      x: inner.x,
      y: inner.y,
      width: inner.width,
      height: inner.height - serialHeight - (hasSerial ? 4 * scale : 0),
    },
    serial: hasSerial
      ? {
          x: inner.x,
          y: inner.y + inner.height - serialHeight,
          width: inner.width,
          height: serialHeight,
        }
      : null,
  };
}

function drawMainTicketLayout(targetCtx, layout, state, scale, orientation) {
  if (orientation === "horizontal") {
    drawMainHorizontalLayout(targetCtx, layout, state, scale);
    return;
  }

  drawMainHeader(targetCtx, layout.title, state, scale);
  drawMainBody(targetCtx, layout.details, state, scale, orientation);
  if (layout.notes) {
    drawMainFooter(targetCtx, layout.notes, state, scale);
  }
}

function drawStubLayout(targetCtx, layout, state, scale, orientation, serial) {
  if (orientation === "horizontal") {
    drawStubHorizontalLayout(targetCtx, layout, state, scale, serial);
    return;
  }

  drawStubHeader(targetCtx, layout.info, state, scale, orientation);
  if (serial && layout.serial) {
    drawSerialBlock(targetCtx, layout.serial, serial, state, scale, orientation);
  }
}

function drawMainHorizontalLayout(targetCtx, layout, state, scale) {
  targetCtx.save();
  targetCtx.fillStyle = state.fontColor || "#111111";
  targetCtx.textAlign = "left";

  const titleSize = mmFontToPx(state.fontSizeEventName, scale, 1.05);
  drawWrappedTextBlock(targetCtx, {
    text: state.eventName || "Event Title",
    x: layout.title.x,
    y: layout.title.y,
    maxWidth: layout.title.width,
    maxLines: 2,
    lineHeight: Math.max(titleSize * 1.08, 6 * scale),
    fontSize: titleSize,
    fontFamily: "'IBM Plex Sans JP', sans-serif",
    fontWeight: 700,
    align: "left",
  });

  drawHorizontalMeta(targetCtx, layout.meta, state, scale);
  drawHorizontalDetails(targetCtx, layout.details, state, scale);

  if (layout.notes) {
    const notesLineHeight = Math.max(mmFontToPx(state.fontSizeNotes, scale) * 1.15, 3.8 * scale);
    const notesLines = Math.min(3, wrapTextByWidth(targetCtx, state.notes, layout.notes.width).length || 1);
    const notesHeight = notesLineHeight * notesLines;
    const notesStartY = layout.notes.y + Math.max(0, layout.notes.height - notesHeight - 1 * scale);

    drawWrappedTextBlock(targetCtx, {
      text: state.notes,
      x: layout.notes.x,
      y: notesStartY,
      maxWidth: layout.notes.width,
      maxLines: 3,
      lineHeight: notesLineHeight,
      fontSize: mmFontToPx(state.fontSizeNotes, scale),
      fontFamily: "'IBM Plex Sans JP', sans-serif",
      fontWeight: 400,
      align: "left",
    });
  }

  targetCtx.restore();
}

function drawHorizontalMeta(targetCtx, rect, state, scale) {
  const metaItems = [
    { label: "DATE", value: state.date, size: state.fontSizeDate },
    { label: "TIME", value: state.time, size: state.fontSizeTime },
    { label: "VENUE", value: state.venue, size: state.fontSizeVenue },
  ].filter((item) => item.value);

  const labelSize = Math.max(6, 2.5 * scale);
  let cursorY = rect.y;
  const itemGap = 2.8 * scale;

  metaItems.forEach((item) => {
    targetCtx.fillStyle = state.labelColor || "#666666";
    targetCtx.font = `500 ${labelSize}px 'IBM Plex Mono', monospace`;
    targetCtx.fillText(item.label, rect.x, cursorY + 3.2 * scale);

    targetCtx.fillStyle = state.fontColor || "#111111";
    drawWrappedTextBlock(targetCtx, {
      text: item.label === "TIME" ? `開演 ${item.value}` : item.value,
      x: rect.x,
      y: cursorY + 3.8 * scale,
      maxWidth: rect.width,
      maxLines: 1,
      lineHeight: Math.max(mmFontToPx(item.size, scale) * 1.1, 4 * scale),
      fontSize: mmFontToPx(item.size, scale),
      fontFamily: "'IBM Plex Sans JP', sans-serif",
      fontWeight: 500,
      align: "left",
    });
    cursorY += Math.max(mmFontToPx(item.size, scale) * 1.25, 5 * scale) + itemGap;
  });
}

function drawHorizontalDetails(targetCtx, rect, state, scale) {
  const items = [
    { label: "SEAT", value: state.seat, size: state.fontSizeSeat },
    { label: "PRICE", value: state.price, size: state.fontSizePrice },
  ].filter((item) => item.value);

  if (!items.length) {
    return;
  }

  const colWidth = rect.width / items.length;
  const labelSize = Math.max(6, 2.5 * scale);
  items.forEach((item, index) => {
    const x = rect.x + colWidth * index;
    targetCtx.fillStyle = state.labelColor || "#666666";
    targetCtx.font = `500 ${labelSize}px 'IBM Plex Mono', monospace`;
    targetCtx.fillText(item.label, x, rect.y + 3 * scale);
    targetCtx.fillStyle = state.fontColor || "#111111";
    drawWrappedTextBlock(targetCtx, {
      text: item.value,
      x,
      y: rect.y + 3.8 * scale,
      maxWidth: colWidth - 3 * scale,
      maxLines: 1,
      lineHeight: Math.max(mmFontToPx(item.size, scale) * 1.12, 4.5 * scale),
      fontSize: mmFontToPx(item.size, scale),
      fontFamily: "'IBM Plex Sans JP', sans-serif",
      fontWeight: 600,
      align: "left",
    });
  });
}

function drawStubHorizontalLayout(targetCtx, layout, state, scale, serial) {
  const inner = layout.info;
  const titleSize = mmFontToPx(state.fontSizeStubEventName, scale, 0.95);
  const metaSize = Math.max(mmFontToPx(state.fontSizeStubDate, scale), mmFontToPx(state.fontSizeStubTime, scale));
  const serialSize = mmFontToPx(state.fontSizeSerial, scale, 1.02);

  targetCtx.save();
  targetCtx.fillStyle = state.fontColor || "#111111";
  targetCtx.textAlign = "left";

  let cursorY = inner.y;
  cursorY += drawWrappedTextBlock(targetCtx, {
    text: state.eventName || "Event Title",
    x: inner.x,
    y: cursorY,
    maxWidth: inner.width,
    maxLines: 3,
    lineHeight: Math.max(titleSize * 1.08, 4.4 * scale),
    fontSize: titleSize,
    fontFamily: "'IBM Plex Sans JP', sans-serif",
    fontWeight: 600,
    align: "left",
  });

  [state.date, state.time ? `開演 ${state.time}` : ""].filter(Boolean).forEach((line) => {
    cursorY += 1.6 * scale;
    cursorY += drawWrappedTextBlock(targetCtx, {
      text: line,
      x: inner.x,
      y: cursorY,
      maxWidth: inner.width,
      maxLines: 1,
      lineHeight: Math.max(metaSize * 1.12, 3.8 * scale),
      fontSize: line.includes("開演") ? mmFontToPx(state.fontSizeStubTime, scale) : mmFontToPx(state.fontSizeStubDate, scale),
      fontFamily: "'IBM Plex Sans JP', sans-serif",
      fontWeight: 400,
      align: "left",
    });
  });

  if (serial) {
    targetCtx.fillStyle = state.labelColor || "#666666";
    targetCtx.font = `500 ${Math.max(7, 2.5 * scale)}px 'IBM Plex Mono', monospace`;
    targetCtx.fillText("No.", layout.serial.x, layout.serial.y + 3 * scale);
    targetCtx.fillStyle = state.fontColor || "#111111";
    targetCtx.font = `600 ${serialSize}px 'IBM Plex Mono', monospace`;
    targetCtx.fillText(serial, layout.serial.x, layout.serial.y + layout.serial.height - 1 * scale);
  }

  targetCtx.restore();
}

function drawMainHeader(targetCtx, rect, state, scale) {
  const titleSize = mmFontToPx(state.fontSizeEventName, scale);
  const dateSize = mmFontToPx(state.fontSizeDate, scale);
  const timeSize = mmFontToPx(state.fontSizeTime, scale);
  const venueSize = mmFontToPx(state.fontSizeVenue, scale);

  targetCtx.save();
  targetCtx.fillStyle = state.fontColor || "#111111";
  targetCtx.textAlign = "center";

  let cursorY = rect.y + 3 * scale;
  cursorY += drawWrappedTextBlock(targetCtx, {
    text: state.eventName || "Event Title",
    x: rect.x + rect.width / 2,
    y: cursorY,
    maxWidth: rect.width - 6 * scale,
    maxLines: 2,
    lineHeight: Math.max(titleSize * 1.1, 5.4 * scale),
    fontSize: titleSize,
    fontFamily: "'IBM Plex Sans JP', sans-serif",
    fontWeight: 700,
    align: "center",
  });

  const metaParts = [state.date, state.time ? `開演 ${state.time}` : "", state.venue].filter(Boolean);
  const metaText = metaParts.join("   ");
  if (metaText) {
    cursorY += 1.6 * scale;
    drawWrappedTextBlock(targetCtx, {
      text: metaText,
      x: rect.x + rect.width / 2,
      y: cursorY,
      maxWidth: rect.width - 8 * scale,
      maxLines: 2,
      lineHeight: Math.max(Math.max(dateSize, timeSize, venueSize) * 1.15, 4.4 * scale),
      fontSize: Math.max(dateSize, timeSize, venueSize),
      fontFamily: "'IBM Plex Sans JP', sans-serif",
      fontWeight: 400,
      align: "center",
    });
  }

  targetCtx.restore();
}

function drawMainBody(targetCtx, rect, state, scale, orientation) {
  const rows = [
    { label: "会場", value: state.venue, size: state.fontSizeVenue },
    { label: "SEAT", value: state.seat, size: state.fontSizeSeat },
    { label: "料金", value: state.price, size: state.fontSizePrice },
  ].filter((row) => row.value);

  if (!rows.length) {
    return;
  }

  targetCtx.save();
  targetCtx.fillStyle = state.fontColor || "#111111";
  const gap = 3.2 * scale;
  const availableHeight = rect.height - gap * (rows.length - 1);
  const rowHeight = Math.max(6 * scale, availableHeight / rows.length);

  rows.forEach((row, index) => {
    const rowRect = {
      x: rect.x,
      y: rect.y + index * (rowHeight + gap),
      width: rect.width,
      height: rowHeight,
    };
    drawInfoRow(targetCtx, rowRect, row, scale, orientation, state);
  });
  targetCtx.restore();
}

function drawInfoRow(targetCtx, rect, row, scale, orientation, state) {
  const labelWidth = Math.min(rect.width * 0.22, 16 * scale);
  const labelX = rect.x;
  const valueX = rect.x + labelWidth + 3 * scale;
  const valueWidth = rect.width - labelWidth - 3 * scale;
  const labelSize = Math.max(6, 2.6 * scale);
  const valueSize = mmFontToPx(row.size, scale);

  targetCtx.textAlign = "left";
  targetCtx.font = `500 ${labelSize}px 'IBM Plex Mono', monospace`;
  targetCtx.fillStyle = state.labelColor || "#666666";
  targetCtx.fillText(row.label, labelX, rect.y + Math.min(rect.height * 0.55, 5.4 * scale));

  targetCtx.fillStyle = state.fontColor || "#111111";
  drawWrappedTextBlock(targetCtx, {
    text: row.value,
    x: valueX,
    y: rect.y,
    maxWidth: valueWidth,
    maxLines: orientation === "vertical" ? 2 : 1,
    lineHeight: Math.max(valueSize * 1.15, 4.8 * scale),
    fontSize: valueSize,
    fontFamily: "'IBM Plex Sans JP', sans-serif",
    fontWeight: 500,
    align: "left",
  });
}

function drawMainFooter(targetCtx, rect, state, scale) {
  const notesSize = mmFontToPx(state.fontSizeNotes, scale);

  targetCtx.save();
  targetCtx.fillStyle = state.labelColor || "#666666";
  targetCtx.textAlign = "left";
  targetCtx.font = `500 ${Math.max(7, 2.7 * scale)}px 'IBM Plex Mono', monospace`;
  targetCtx.fillText("NOTE", rect.x, rect.y + 3.6 * scale);

  targetCtx.fillStyle = state.fontColor || "#111111";
  drawWrappedTextBlock(targetCtx, {
    text: state.notes,
    x: rect.x,
    y: rect.y + 4.5 * scale,
    maxWidth: rect.width,
    maxLines: 3,
    lineHeight: Math.max(notesSize * 1.18, 4 * scale),
    fontSize: notesSize,
    fontFamily: "'IBM Plex Sans JP', sans-serif",
    fontWeight: 400,
    align: "left",
  });
  targetCtx.restore();
}

function drawStubHeader(targetCtx, rect, state, scale, orientation) {
  const titleSize = mmFontToPx(state.fontSizeStubEventName, scale);
  const dateSize = mmFontToPx(state.fontSizeStubDate, scale);
  const timeSize = mmFontToPx(state.fontSizeStubTime, scale);

  targetCtx.save();
  targetCtx.fillStyle = state.fontColor || "#111111";
  targetCtx.textAlign = "left";

  let cursorY = rect.y + 1 * scale;
  cursorY += drawWrappedTextBlock(targetCtx, {
    text: state.eventName || "Event Title",
    x: rect.x,
    y: cursorY,
    maxWidth: rect.width,
    maxLines: orientation === "horizontal" ? 3 : 2,
    lineHeight: Math.max(titleSize * 1.08, 4.2 * scale),
    fontSize: titleSize,
    fontFamily: "'IBM Plex Sans JP', sans-serif",
    fontWeight: 600,
    align: "left",
  });

  const metaLines = [state.date, state.time ? `開演 ${state.time}` : ""].filter(Boolean);
  metaLines.forEach((line) => {
    cursorY += 1.6 * scale;
    cursorY += drawWrappedTextBlock(targetCtx, {
      text: line,
      x: rect.x,
      y: cursorY,
      maxWidth: rect.width,
      maxLines: 1,
      lineHeight: Math.max(Math.max(dateSize, timeSize) * 1.15, 3.8 * scale),
      fontSize: line.includes("開演") ? timeSize : dateSize,
      fontFamily: "'IBM Plex Sans JP', sans-serif",
      fontWeight: 400,
      align: "left",
    });
  });

  targetCtx.restore();
}

function drawSerialBlock(targetCtx, rect, serial, state, scale, orientation) {
  const serialSize = mmFontToPx(state.fontSizeSerial, scale);
  targetCtx.save();
  if (orientation === "horizontal") {
    targetCtx.fillStyle = state.labelColor || "#666666";
    targetCtx.textAlign = "left";
    targetCtx.font = `500 ${Math.max(7, 2.7 * scale)}px 'IBM Plex Mono', monospace`;
    targetCtx.fillText("No.", rect.x, rect.y + 3.4 * scale);

    targetCtx.fillStyle = state.fontColor || "#111111";
    targetCtx.textAlign = "left";
    targetCtx.font = `600 ${serialSize}px 'IBM Plex Mono', monospace`;
    targetCtx.fillText(serial, rect.x, rect.y + rect.height - 1.4 * scale);
  } else {
    const blockText = `No. ${serial}`;
    targetCtx.fillStyle = state.fontColor || "#111111";
    targetCtx.textAlign = "right";
    targetCtx.font = `600 ${serialSize}px 'IBM Plex Mono', monospace`;
    targetCtx.fillText(blockText, rect.x + rect.width, rect.y + rect.height - 1.8 * scale);
  }
  targetCtx.restore();
}

function insetRect(rect, insetX, insetY = insetX) {
  return {
    x: rect.x + insetX,
    y: rect.y + insetY,
    width: rect.width - insetX * 2,
    height: rect.height - insetY * 2,
  };
}

function toNumber(value, fallback) {
  return Number.isFinite(Number(value)) ? Number(value) : fallback;
}

function toInteger(value, fallback) {
  return Number.isInteger(Number(value)) ? Number(value) : fallback;
}

function normalizeEnum(value, allowedValues, fallback) {
  return allowedValues.includes(value) ? value : fallback;
}

function isBetween(value, min, max) {
  return Number.isFinite(value) && value >= min && value <= max;
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function mmFontToPx(ptValue, scale, ratio = 1) {
  return Math.max(5, ptValue * scale * 0.28 * ratio);
}

function normalizeSerialMode(state) {
  if (state.serialMode) {
    return state.serialMode;
  }
  return /^#+$/.test(state.serialStart || "") ? "placeholder" : "numbered";
}

function getSerialModeHint(mode) {
  return mode === "placeholder"
    ? "`###` をプレースホルダーとしてPDFに残し、印刷所で連番差し込みを依頼します。"
    : "`001` や `015` のような開始番号を入れると、このアプリで連番付きPDFを複数ページ生成します。";
}

function parseSerialStart(startRaw) {
  const match = startRaw.match(/^(\D*)(\d+)(\D*)$/);
  if (!match) {
    return null;
  }

  const [, prefix, digitsRaw, suffix] = match;
  return {
    prefix,
    suffix,
    digits: digitsRaw.length,
    startNumber: Number(digitsRaw),
  };
}

function resolveSerialSettings(state) {
  if (!state.serialEnabled) {
    return {
      enabled: false,
      mode: "off",
      isValid: true,
      pageCount: 1,
      previewText: "",
    };
  }

  const mode = normalizeSerialMode(state);
  const startRaw = state.serialStart || (mode === "placeholder" ? "###" : "001");
  const count = Number.isInteger(state.serialCount) ? state.serialCount : 1;

  if (mode === "placeholder") {
    if (!/^#+$/.test(startRaw)) {
      return {
        enabled: true,
        mode,
        isValid: false,
        errorMessage: "差し込み記号は `###` のように # のみで入力してください。",
      };
    }

    const digits = startRaw.length;
    const first = String(1).padStart(digits, "0");
    const last = String(1 + Math.max(0, count - 1)).padStart(digits, "0");
    return {
      enabled: true,
      mode,
      isValid: true,
      pageCount: 1,
      previewText: startRaw,
      placeholder: startRaw,
      startDisplay: first,
      endDisplay: last,
      digits,
      format: first,
    };
  }

  const parsed = parseSerialStart(startRaw);
  if (!parsed) {
    return {
      enabled: true,
      mode,
      isValid: false,
      errorMessage: "開始番号は `001` や `A001` のように数字を含む形式で入力してください。",
    };
  }

  const first = `${parsed.prefix}${String(parsed.startNumber).padStart(parsed.digits, "0")}${parsed.suffix}`;
  const last = `${parsed.prefix}${String(parsed.startNumber + Math.max(0, count - 1)).padStart(parsed.digits, "0")}${parsed.suffix}`;
  return {
    enabled: true,
    mode,
    isValid: true,
    pageCount: count,
    previewText: first,
    placeholder: startRaw,
    startDisplay: first,
    endDisplay: last,
    digits: parsed.digits,
    format: first,
    parsed,
  };
}

function getSerialForIndex(state, index) {
  const serialConfig = resolveSerialSettings(state);
  if (!serialConfig.enabled || !serialConfig.isValid) {
    return "";
  }

  if (serialConfig.mode === "placeholder") {
    return serialConfig.previewText;
  }

  const { prefix, suffix, digits, startNumber } = serialConfig.parsed;
  return `${prefix}${String(startNumber + index).padStart(digits, "0")}${suffix}`;
}

function getSerialPrintComment(state) {
  const serialConfig = resolveSerialSettings(state);
  if (!serialConfig.enabled) {
    return "連番設定がオフです。必要な場合は連番をオンにしてください。";
  }
  if (!serialConfig.isValid) {
    return serialConfig.errorMessage;
  }

  return [
    `チケットの「${serialConfig.placeholder}」部分に整理番号を連番で挿入してください。`,
    "",
    `開始番号：${serialConfig.startDisplay}`,
    `終了番号：${serialConfig.endDisplay}`,
    `桁数：${serialConfig.digits}桁（ゼロ埋め）`,
    "",
    "表示形式：",
    serialConfig.format,
  ].join("\n");
}

async function generatePdf() {
  const state = getFormState();
  const errors = validateState(state);
  renderErrors(errors);

  if (errors.length) {
    return;
  }

  ui.generateButton.disabled = true;

  try {
    const { PDFDocument } = PDFLib;
    const pdfDoc = await PDFDocument.create();
    const ticket = getTicketDimensions(state);
    const pageWidth = mmToPt(ticket.width + BLEED_MM * 2);
    const pageHeight = mmToPt(ticket.height + BLEED_MM * 2);
    const serialConfig = resolveSerialSettings(state);
    const totalPages = serialConfig.enabled && serialConfig.isValid ? serialConfig.pageCount : 1;

    for (let index = 0; index < totalPages; index += 1) {
      const pageCanvas = renderTicketCanvasForPdf(state, getSerialForIndex(state, index), ticket);
      const pngDataUrl = pageCanvas.toDataURL("image/png");
      const imageBytes = dataUrlToUint8Array(pngDataUrl);
      const pngImage = await pdfDoc.embedPng(imageBytes);
      const page = pdfDoc.addPage([pageWidth, pageHeight]);
      page.drawImage(pngImage, {
        x: 0,
        y: 0,
        width: pageWidth,
        height: pageHeight,
      });
    }

    const bytes = await pdfDoc.save();
    downloadPdf(bytes, "ticket.pdf");
  } catch (error) {
    renderErrors(["PDF生成中にエラーが発生しました。入力内容を確認して再度お試しください。"]);
    console.error(error);
  } finally {
    updateGenerateButtonState(!ui.errorBox.classList.contains("is-visible"));
  }
}

function renderTicketCanvasForPdf(state, serial, ticket) {
  const renderScale = 12;
  const canvas = document.createElement("canvas");
  const totalWidth = (ticket.width + BLEED_MM * 2) * renderScale;
  const totalHeight = (ticket.height + BLEED_MM * 2) * renderScale;
  canvas.width = Math.round(totalWidth);
  canvas.height = Math.round(totalHeight);

  const pdfCtx = canvas.getContext("2d");
  const geometry = {
    cssWidth: canvas.width,
    cssHeight: canvas.height,
    scale: renderScale,
    originX: 0,
    originY: 0,
    bleedPx: BLEED_MM * renderScale,
    trimX: BLEED_MM * renderScale,
    trimY: BLEED_MM * renderScale,
    trimWidth: ticket.width * renderScale,
    trimHeight: ticket.height * renderScale,
    bleedWidth: totalWidth,
    bleedHeight: totalHeight,
    ticket,
  };

  pdfCtx.fillStyle = "#ffffff";
  pdfCtx.fillRect(0, 0, canvas.width, canvas.height);
  drawBleedArea(geometry, pdfCtx);
  fillTicketTexture(geometry, state, pdfCtx);
  drawTrimLine(geometry, pdfCtx);
  drawCropMarks(geometry, pdfCtx);
  drawPerforation(geometry, state, pdfCtx);
  drawTicketTextCanvas(geometry, state, serial, pdfCtx);

  return canvas;
}

function mmToPt(mm) {
  return (mm / 25.4) * 72;
}

function dataUrlToUint8Array(dataUrl) {
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function downloadPdf(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
