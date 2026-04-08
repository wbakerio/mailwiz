const MAX_ELEMENTS = 7;
const STORAGE_KEY = "mail-calculator-saved-items";
const SETTINGS_STORAGE_KEY = "mail-calculator-settings";
const FOLD_MULTIPLIERS = {
  none: 1,
  bifold: 2,
  trifold: 3,
  quadfold: 4
};
const DEFAULT_SETTINGS = {
  weightUnit: "oz",
  thicknessUnit: "in",
  fudgeFactor: 1.6
};

const PAPER_TYPES = {
  text: {
    label: "Text / Book / Offset",
    basisWidth: 25,
    basisHeight: 38
  },
  bond: {
    label: "Bond",
    basisWidth: 17,
    basisHeight: 22
  },
  c1s: {
    label: "C1S",
    basisWidth: 20,
    basisHeight: 26
  },
  c2s: {
    label: "C2S",
    basisWidth: 20,
    basisHeight: 26
  },
  cover: {
    label: "Cover",
    basisWidth: 20,
    basisHeight: 26
  },
  bristol: {
    label: "Bristol",
    basisWidth: 22.5,
    basisHeight: 28.5
  },
  index: {
    label: "Index",
    basisWidth: 25.5,
    basisHeight: 30.5
  },
  tag: {
    label: "Tag",
    basisWidth: 24,
    basisHeight: 36
  }
};

const STOCK_LIBRARY = [
  { type: "cover", finish: "Gloss", weight: "80", weightLabel: "80", thicknessIn: 0.0077 },
  { type: "cover", finish: "Gloss", weight: "100", weightLabel: "100", thicknessIn: 0.0099 },
  { type: "cover", finish: "Gloss", weight: "110", weightLabel: "110", thicknessIn: 0.0112 },
  { type: "cover", finish: "Gloss", weight: "120", weightLabel: "120", thicknessIn: 0.0120 },
  { type: "cover", finish: "Gloss", weight: "130", weightLabel: "130", thicknessIn: 0.0130 },
  { type: "cover", finish: "Dull", weight: "80", weightLabel: "80", thicknessIn: 0.0090 },
  { type: "cover", finish: "Dull", weight: "100", weightLabel: "100", thicknessIn: 0.0111 },
  { type: "cover", finish: "Dull", weight: "110", weightLabel: "110", thicknessIn: 0.0126 },
  { type: "cover", finish: "Dull", weight: "120", weightLabel: "120", thicknessIn: 0.0130 },
  { type: "cover", finish: "Dull", weight: "130", weightLabel: "130", thicknessIn: 0.0147 },
  { type: "cover", finish: "Smooth", weight: "65", weightLabel: "65", thicknessIn: 0.0088 },
  { type: "cover", finish: "Smooth", weight: "80", weightLabel: "80", thicknessIn: 0.0105 },
  { type: "cover", finish: "Smooth", weight: "100", weightLabel: "100", thicknessIn: 0.0130 },
  { type: "cover", finish: "Smooth", weight: "110", weightLabel: "110", thicknessIn: 0.0145 },
  { type: "cover", finish: "Smooth", weight: "130", weightLabel: "130", thicknessIn: 0.0175 },
  { type: "cover", finish: "Smooth", weight: "160", weightLabel: "160", thicknessIn: 0.0218 },
  { type: "cover", finish: "Linen", weight: "80", weightLabel: "80", thicknessIn: 0.0110 },
  { type: "cover", finish: "Linen", weight: "100", weightLabel: "100", thicknessIn: 0.0125 },
  { type: "cover", finish: "Linen", weight: "120", weightLabel: "120", thicknessIn: 0.0155 },
  { type: "cover", finish: "Linen", weight: "130", weightLabel: "130", thicknessIn: 0.0170 },
  { type: "cover", finish: "Laid", weight: "80", weightLabel: "80", thicknessIn: 0.0135 },
  { type: "cover", finish: "Laid", weight: "100", weightLabel: "100", thicknessIn: 0.0170 },
  { type: "cover", finish: "Laid", weight: "120", weightLabel: "120", thicknessIn: 0.0190 },
  { type: "bond", finish: "Smooth", weight: "20", weightLabel: "20", thicknessIn: 0.0040 },
  { type: "bond", finish: "Smooth", weight: "24", weightLabel: "24", thicknessIn: 0.0045 },
  { type: "bond", finish: "Smooth", weight: "28", weightLabel: "28", thicknessIn: 0.0044 },
  { type: "bond", finish: "Smooth", weight: "32", weightLabel: "32", thicknessIn: 0.0047 },
  { type: "bond", finish: "Linen", weight: "24", weightLabel: "24", thicknessIn: 0.0045 },
  { type: "bond", finish: "Laid", weight: "24", weightLabel: "24", thicknessIn: 0.0050 },
  { type: "text", finish: "Gloss", weight: "70", weightLabel: "70", thicknessIn: 0.0033 },
  { type: "text", finish: "Gloss", weight: "80", weightLabel: "80", thicknessIn: 0.0037 },
  { type: "text", finish: "Gloss", weight: "100", weightLabel: "100", thicknessIn: 0.0047 },
  { type: "text", finish: "Dull", weight: "70", weightLabel: "70", thicknessIn: 0.0037 },
  { type: "text", finish: "Dull", weight: "80", weightLabel: "80", thicknessIn: 0.0042 },
  { type: "text", finish: "Dull", weight: "100", weightLabel: "100", thicknessIn: 0.0057 },
  { type: "text", finish: "Smooth", weight: "40", weightLabel: "40", thicknessIn: 0.0031 },
  { type: "text", finish: "Smooth", weight: "50", weightLabel: "50", thicknessIn: 0.0039 },
  { type: "text", finish: "Smooth", weight: "60", weightLabel: "60", thicknessIn: 0.0042 },
  { type: "text", finish: "Smooth", weight: "70", weightLabel: "70", thicknessIn: 0.0047 },
  { type: "text", finish: "Smooth", weight: "80", weightLabel: "80", thicknessIn: 0.0054 },
  { type: "text", finish: "Smooth", weight: "100", weightLabel: "100", thicknessIn: 0.0066 },
  { type: "text", finish: "Linen", weight: "70", weightLabel: "70", thicknessIn: 0.0048 },
  { type: "text", finish: "Linen", weight: "80", weightLabel: "80", thicknessIn: 0.0056 },
  { type: "text", finish: "Linen", weight: "100", weightLabel: "100", thicknessIn: 0.0069 },
  { type: "bristol", finish: "Vellum", weight: "67", weightLabel: "67", thicknessIn: 0.0093 },
  { type: "index", finish: "Smooth", weight: "90", weightLabel: "90", thicknessIn: 0.0074 },
  { type: "index", finish: "Smooth", weight: "110", weightLabel: "110", thicknessIn: 0.0095 },
  { type: "tag", finish: "Smooth", weight: "125", weightLabel: "125", thicknessIn: 0.0095 },
  { type: "tag", finish: "Smooth", weight: "150", weightLabel: "150", thicknessIn: 0.0116 },
  { type: "c1s", finish: "C1S", weight: "8", weightLabel: "8pt", thicknessIn: 0.008, gsm: 180, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "10", weightLabel: "10pt", thicknessIn: 0.010, gsm: 203, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "12", weightLabel: "12pt", thicknessIn: 0.012, gsm: 231, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "14", weightLabel: "14pt", thicknessIn: 0.014, gsm: 257, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "15", weightLabel: "15pt", thicknessIn: 0.015, gsm: 275, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "16", weightLabel: "16pt", thicknessIn: 0.016, gsm: 292, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "18", weightLabel: "18pt", thicknessIn: 0.018, gsm: 320, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "20", weightLabel: "20pt", thicknessIn: 0.020, gsm: 354, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "24", weightLabel: "24pt", thicknessIn: 0.024, gsm: 408, pointBased: true },
  { type: "c1s", finish: "C1S", weight: "28", weightLabel: "28pt", thicknessIn: 0.028, gsm: 480, pointBased: true },
  { type: "c2s", finish: "C2S", weight: "8", weightLabel: "8pt", thicknessIn: 0.008, gsm: 195, pointBased: true },
  { type: "c2s", finish: "C2S", weight: "10", weightLabel: "10pt", thicknessIn: 0.010, gsm: 254, pointBased: true },
  { type: "c2s", finish: "C2S", weight: "12", weightLabel: "12pt", thicknessIn: 0.012, gsm: 278, pointBased: true },
  { type: "c2s", finish: "C2S", weight: "14", weightLabel: "14pt", thicknessIn: 0.014, gsm: 319, pointBased: true },
  { type: "c2s", finish: "C2S", weight: "16", weightLabel: "16pt", thicknessIn: 0.016, gsm: 353, pointBased: true },
  { type: "c2s", finish: "C2S", weight: "18", weightLabel: "18pt", thicknessIn: 0.018, gsm: 358, pointBased: true },
  { type: "c2s", finish: "C2S", weight: "24", weightLabel: "24pt", thicknessIn: 0.024, gsm: 465, pointBased: true }
];

const defaultSavedItems = [
  { id: createId(), name: "#9 Remit Envelope", weightOz: 0.22, thicknessIn: 0.024 },
  { id: createId(), name: "#10 Reg - Diagonal Seam", weightOz: 0.1825, thicknessIn: 0.024 },
  { id: createId(), name: "#10 Window - Diagonal Seam", weightOz: 0.1825, thicknessIn: 0.024 },
  { id: createId(), name: "#10 Window - Side Seam", weightOz: 0.1875, thicknessIn: 0.024 },
  { id: createId(), name: "#10 Reg - Side Seam", weightOz: 0.1875, thicknessIn: 0.024 },
  { id: createId(), name: "A6 Envelope", weightOz: 0.15375, thicknessIn: 0.024 },
  { id: createId(), name: "A7 Envelope", weightOz: 0.18125, thicknessIn: 0.024 },
  { id: createId(), name: "A2 Envelope", weightOz: 0.125, thicknessIn: 0.024 },
  { id: createId(), name: "A7 - Machine Insertable Envelope", weightOz: 0.1975, thicknessIn: 0.024 },
  { id: createId(), name: "#14 Window Envelope", weightOz: 0.2725, thicknessIn: 0.024 },
  { id: createId(), name: "24lb - 6x9 Catalog Envelope", weightOz: 0.2675, thicknessIn: 0.024 },
  { id: createId(), name: "7.5x10.5 Catalog Envelope", weightOz: 0.38, thicknessIn: 0.024 },
  { id: createId(), name: "28lb - 9x12 Booklet Envelope", weightOz: 0.61, thicknessIn: 0.025 },
  { id: createId(), name: "24lb - 6x9 Booklet Envelope", weightOz: 0.2525, thicknessIn: 0.024 },
  { id: createId(), name: "28lb - 9x12 Catalog Envelope", weightOz: 0.5875, thicknessIn: 0.025 }
];

const state = {
  elements: [],
  savedItems: loadSavedItems(),
  settings: loadSettings()
};

const elementsRoot = document.querySelector("#elements");
const addElementBtn = document.querySelector("#add-element-btn");
const clearBuildBtn = document.querySelector("#clear-build-btn");
const elementTemplate = document.querySelector("#element-template");
const totalWeightEl = document.querySelector("#total-weight");
const totalThicknessEl = document.querySelector("#total-thickness");
const totalElementsEl = document.querySelector("#total-elements");
const totalWeightLabelEl = document.querySelector("#total-weight-label");
const totalThicknessLabelEl = document.querySelector("#total-thickness-label");
const settingsBtn = document.querySelector("#settings-btn");
const settingsManager = document.querySelector("#settings-manager");
const closeSettingsManagerBtn = document.querySelector("#close-settings-manager-btn");
const settingsForm = document.querySelector("#settings-form");
const weightUnitInput = document.querySelector("#weight-unit-input");
const thicknessUnitInput = document.querySelector("#thickness-unit-input");
const fudgeFactorInput = document.querySelector("#fudge-factor-input");
const manageSavedBtn = document.querySelector("#manage-saved-btn");
const savedManager = document.querySelector("#saved-manager");
const closeSavedManagerBtn = document.querySelector("#close-saved-manager-btn");
const customItemForm = document.querySelector("#custom-item-form");
const savedItemsRoot = document.querySelector("#saved-items");
const saveCustomItemBtn = document.querySelector("#save-custom-item-btn");
const cancelEditBtn = document.querySelector("#cancel-edit-btn");
const customWeightLabelEl = document.querySelector("#custom-weight-label");
const customThicknessLabelEl = document.querySelector("#custom-thickness-label");

let editingSavedItemId = null;

function loadSavedItems() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);

    if (!raw) {
      return defaultSavedItems;
    }

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return defaultSavedItems;
    }

    const validParsed = parsed.filter((item) => item && item.id && item.name);
    const seededNames = new Set(defaultSavedItems.map((item) => item.name));
    const customItems = validParsed.filter((item) => !seededNames.has(item.name));

    return [...defaultSavedItems, ...customItems];
  } catch {
    return defaultSavedItems;
  }
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) {
      return { ...DEFAULT_SETTINGS };
    }

    const parsed = JSON.parse(raw);
    return {
      weightUnit: ["oz", "lb", "g"].includes(parsed.weightUnit) ? parsed.weightUnit : DEFAULT_SETTINGS.weightUnit,
      thicknessUnit: ["in", "mm"].includes(parsed.thicknessUnit) ? parsed.thicknessUnit : DEFAULT_SETTINGS.thicknessUnit,
      fudgeFactor: Number(parsed.fudgeFactor) > 0 ? Number(parsed.fudgeFactor) : DEFAULT_SETTINGS.fudgeFactor
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

function persistSavedItems() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.savedItems));
}

function persistSettings() {
  localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(state.settings));
}

function openSettingsManager() {
  weightUnitInput.value = state.settings.weightUnit;
  thicknessUnitInput.value = state.settings.thicknessUnit;
  fudgeFactorInput.value = String(state.settings.fudgeFactor);
  settingsManager.classList.remove("hidden");
  settingsManager.setAttribute("aria-hidden", "false");
}

function closeSettingsManager() {
  settingsManager.classList.add("hidden");
  settingsManager.setAttribute("aria-hidden", "true");
}

function openSavedManager() {
  savedManager.classList.remove("hidden");
  savedManager.setAttribute("aria-hidden", "false");
}

function closeSavedManager() {
  savedManager.classList.add("hidden");
  savedManager.setAttribute("aria-hidden", "true");
  clearSavedItemForm();
}

function startEditingSavedItem(item) {
  editingSavedItemId = item.id;
  customItemForm.name.value = item.name;
  customItemForm.weight.value = convertWeightFromOz(item.weightOz);
  customItemForm.thickness.value = convertThicknessFromIn(item.thicknessIn);
  saveCustomItemBtn.textContent = "Update Item";
  cancelEditBtn.classList.remove("hidden");
  openSavedManager();
}

function clearSavedItemForm() {
  editingSavedItemId = null;
  customItemForm.reset();
  saveCustomItemBtn.textContent = "Save Item";
  cancelEditBtn.classList.add("hidden");
}

function createId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }

  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function convertWeightFromOz(valueOz) {
  switch (state.settings.weightUnit) {
    case "lb":
      return valueOz / 16;
    case "g":
      return valueOz * 28.349523125;
    default:
      return valueOz;
  }
}

function convertWeightToOz(value) {
  switch (state.settings.weightUnit) {
    case "lb":
      return value * 16;
    case "g":
      return value / 28.349523125;
    default:
      return value;
  }
}

function convertThicknessFromIn(valueIn) {
  return state.settings.thicknessUnit === "mm" ? valueIn * 25.4 : valueIn;
}

function convertThicknessToIn(value) {
  return state.settings.thicknessUnit === "mm" ? value / 25.4 : value;
}

function createElement() {
  return {
    id: createId(),
    mode: "paper",
    isCollapsed: false,
    quantity: 1,
    paperWeight: "",
    paperType: "",
    paperFinish: "",
    foldType: "none",
    width: "",
    length: "",
    savedItemId: state.savedItems[0]?.id ?? "",
    result: { weightOz: 0, thicknessIn: 0, poundsPerSheet: 0, gsm: 0 }
  };
}

function getWeightOptionLabel(paperTypeKey, value) {
  const stock = STOCK_LIBRARY.find((item) => item.type === paperTypeKey && item.weight === String(value));
  return stock?.weightLabel ?? String(value);
}

function getMatchingStockOptions(selections, ignoredField = "") {
  return STOCK_LIBRARY.filter((item) => {
    if (ignoredField !== "paperType" && selections.paperType && item.type !== selections.paperType) {
      return false;
    }

    if (ignoredField !== "paperFinish" && selections.paperFinish && item.finish !== selections.paperFinish) {
      return false;
    }

    if (ignoredField !== "paperWeight" && selections.paperWeight && item.weight !== String(selections.paperWeight)) {
      return false;
    }

    return true;
  });
}

function populateSelect(select, placeholderText, options, currentValue = "") {
  select.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = placeholderText;
  select.append(placeholder);

  options.forEach((optionData) => {
    const option = document.createElement("option");
    option.value = optionData.value;
    option.textContent = optionData.label;
    select.append(option);
  });

  const hasCurrent = options.some((option) => option.value === String(currentValue));
  select.value = hasCurrent ? String(currentValue) : "";
  select.disabled = options.length === 0;
}

function syncPaperStockOptions(paperTypeInput, paperFinishInput, paperWeightInput, element) {
  const selections = {
    paperType: paperTypeInput.value,
    paperFinish: paperFinishInput.value,
    paperWeight: paperWeightInput.value
  };

  const availableTypes = [...new Map(
    getMatchingStockOptions(selections, "paperType")
      .map((item) => [item.type, { value: item.type, label: PAPER_TYPES[item.type].label }])
  ).values()];

  const availableFinishes = [...new Map(
    getMatchingStockOptions(selections, "paperFinish")
      .map((item) => [item.finish, { value: item.finish, label: item.finish }])
  ).values()];

  const availableWeights = [...new Map(
    getMatchingStockOptions(selections, "paperWeight")
      .map((item) => [item.weight, { value: item.weight, label: item.weightLabel }])
  ).values()];

  populateSelect(paperTypeInput, "Choose paper type", availableTypes, selections.paperType);
  populateSelect(paperFinishInput, "Choose finish", availableFinishes, selections.paperFinish);
  populateSelect(paperWeightInput, "Choose paper weight", availableWeights, selections.paperWeight);

  element.paperType = paperTypeInput.value;
  element.paperFinish = paperFinishInput.value;
  element.paperWeight = paperWeightInput.value;
}

function calculatePaperMetrics(element) {
  const paperWeight = Number(element.paperWeight);
  const width = Number(element.width);
  const length = Number(element.length);
  const quantity = Math.max(1, Number(element.quantity) || 1);
  const paperType = PAPER_TYPES[element.paperType];
  const foldMultiplier = FOLD_MULTIPLIERS[element.foldType] ?? 1;
  const foldFudgeFactor = element.foldType === "none" ? 1 : state.settings.fudgeFactor;
  const stock = STOCK_LIBRARY.find((item) =>
    item.type === element.paperType &&
    item.finish === element.paperFinish &&
    item.weight === String(element.paperWeight)
  );

  if (!paperType || !stock || !paperWeight || !width || !length) {
    return {
      weightOz: 0,
      thicknessIn: 0,
      poundsPerSheet: 0,
      sheetAreaInSq: 0,
      basisAreaInSq: 0,
      reamAreaInSq: 0,
      poundsPerSquareIn: 0,
      basisSizeLabel: "",
      gsm: 0
    };
  }

  const sheetAreaInSq = width * length;
  const basisAreaInSq = paperType.basisWidth * paperType.basisHeight;
  const reamAreaInSq = basisAreaInSq * 500;
  const sheetAreaSquareMeters = sheetAreaInSq * 0.00064516;
  const gsm = stock.gsm ?? ((paperWeight * 453.59237) / (reamAreaInSq * 0.00064516));
  const poundsPerSheet = stock.pointBased
    ? ((gsm * sheetAreaSquareMeters) / 453.59237)
    : paperWeight * (sheetAreaInSq / basisAreaInSq) / 500;
  const poundsPerSquareIn = sheetAreaInSq ? poundsPerSheet / sheetAreaInSq : 0;
  const weightOz = poundsPerSheet * 16 * quantity;
  const thicknessIn = stock.thicknessIn * quantity * foldMultiplier * foldFudgeFactor;

  return {
    weightOz,
    thicknessIn,
    poundsPerSheet,
    sheetAreaInSq,
    basisAreaInSq,
    reamAreaInSq,
    poundsPerSquareIn,
    basisSizeLabel: `${paperType.basisWidth}" x ${paperType.basisHeight}"`,
    gsm,
    finishLabel: stock.finish,
    weightLabel: stock.weightLabel
  };
}

function calculateSavedMetrics(element) {
  const quantity = Math.max(1, Number(element.quantity) || 1);
  const saved = state.savedItems.find((item) => item.id === element.savedItemId);

  if (!saved) {
    return { weightOz: 0, thicknessIn: 0 };
  }

  return {
    weightOz: saved.weightOz * quantity,
    thicknessIn: saved.thicknessIn * quantity,
    poundsPerSheet: saved.weightOz / 16
  };
}

function calculateElement(element) {
  return element.mode === "saved"
    ? calculateSavedMetrics(element)
    : calculatePaperMetrics(element);
}

function formatWeight(valueOz) {
  const converted = convertWeightFromOz(valueOz);

  switch (state.settings.weightUnit) {
    case "lb":
      return `${converted.toFixed(4)} lb`;
    case "g":
      return `${converted.toFixed(2)} g`;
    default:
      return `${converted.toFixed(3)} oz`;
  }
}

function formatThickness(valueIn) {
  const converted = convertThicknessFromIn(valueIn);
  return state.settings.thicknessUnit === "mm"
    ? `${converted.toFixed(3)} mm`
    : `${converted.toFixed(4)} in`;
}

function formatWeightPerSquareIn(valueOzPerSquareIn) {
  const converted = convertWeightFromOz(valueOzPerSquareIn);

  switch (state.settings.weightUnit) {
    case "lb":
      return `${converted.toFixed(6)} lb/sq in`;
    case "g":
      return `${converted.toFixed(4)} g/sq in`;
    default:
      return `${converted.toFixed(4)} oz/sq in`;
  }
}

function formatSquareInches(value) {
  return `${value.toLocaleString(undefined, { maximumFractionDigits: 2 })} sq in`;
}

function formatGsm(value) {
  return `${Number(value).toFixed(value % 1 === 0 ? 0 : 1)} gsm`;
}

function formatDimension(value) {
  return value ? `${value}"` : "?";
}

function getFoldLabel(foldType) {
  const labels = {
    none: "No fold",
    bifold: "Bifold",
    trifold: "Trifold",
    quadfold: "Quadfold"
  };

  return labels[foldType] ?? "No fold";
}

function getElementSummary(element) {
  if (element.mode === "saved") {
    const saved = state.savedItems.find((item) => item.id === element.savedItemId);
    const name = saved?.name ?? "Saved item";
    return {
      title: `${name} x${element.quantity || 1}`,
      metrics: `${formatWeight(element.result.weightOz || 0)} | ${formatThickness(element.result.thicknessIn || 0)}`
    };
  }

  const paperType = PAPER_TYPES[element.paperType];
  const paperTypeLabel = paperType?.label ?? "Paper Sheet**";
  const finishLabel = element.paperFinish || "No finish";
  const paperWeightLabel = element.paperWeight
    ? getWeightOptionLabel(element.paperType, element.paperWeight)
    : "No weight";
  const sizeLabel = `${formatDimension(element.width)} x ${formatDimension(element.length)}`;

  return {
    title: `${paperTypeLabel} | ${finishLabel} | ${paperWeightLabel} | ${sizeLabel} | ${getFoldLabel(element.foldType)} | Qty ${element.quantity || 1}`,
    metrics: `${formatWeight(element.result.weightOz || 0)} | ${formatThickness(element.result.thicknessIn || 0)}`
  };
}

function syncSelectPlaceholderState(select) {
  select.classList.toggle("placeholder", !select.value);
}

function getWeightUnitLabel() {
  switch (state.settings.weightUnit) {
    case "lb":
      return "lb";
    case "g":
      return "g";
    default:
      return "oz";
  }
}

function getThicknessUnitLabel() {
  return state.settings.thicknessUnit === "mm" ? "mm" : "in";
}

function refreshAllElementDisplays() {
  document.querySelectorAll(".element-card").forEach((card) => {
    const weightLabel = card.querySelector(".weight-label");
    const thicknessLabel = card.querySelector(".thickness-label");

    if (weightLabel) {
      weightLabel.textContent = `Weight (${getWeightUnitLabel()})`;
    }

    if (thicknessLabel) {
      thicknessLabel.textContent = `Thickness (${getThicknessUnitLabel()})`;
    }
  });
}

function applyDisplaySettings() {
  totalWeightLabelEl.textContent = `Total Weight (${getWeightUnitLabel()})`;
  totalThicknessLabelEl.textContent = `Total Thickness (${getThicknessUnitLabel()})`;
  customWeightLabelEl.textContent = `Weight (${getWeightUnitLabel()})`;
  customThicknessLabelEl.textContent = `Thickness (${getThicknessUnitLabel()})`;
  refreshAllElementDisplays();
  renderSavedItems();
  document.querySelectorAll(".element-card .quantity-input").forEach((input) => {
    input.dispatchEvent(new Event("change"));
  });
  updateTotals();
}

function updateTotals() {
  const totals = state.elements.reduce((accumulator, element) => {
    const result = calculateElement(element);
    element.result = result;
    accumulator.weightOz += result.weightOz;
    accumulator.thicknessIn += result.thicknessIn;
    return accumulator;
  }, { weightOz: 0, thicknessIn: 0 });

  totalWeightEl.textContent = formatWeight(totals.weightOz);
  totalThicknessEl.textContent = formatThickness(totals.thicknessIn);
  totalElementsEl.textContent = String(state.elements.length);

  addElementBtn.disabled = state.elements.length >= MAX_ELEMENTS;
  addElementBtn.textContent = state.elements.length >= MAX_ELEMENTS
    ? "Element Limit Reached"
    : "+ Add Element";
}

function refreshSavedItemSelects() {
  const selects = document.querySelectorAll(".saved-item-input");

  selects.forEach((select) => {
    const currentValue = select.value;
    select.innerHTML = "";

    state.savedItems.forEach((item) => {
      const option = document.createElement("option");
      option.value = item.id;
      option.textContent = item.name;
      select.append(option);
    });

    const hasCurrent = state.savedItems.some((item) => item.id === currentValue);
    select.value = hasCurrent ? currentValue : (state.savedItems[0]?.id ?? "");
  });
}

function renderSavedItems() {
  savedItemsRoot.innerHTML = "";

  state.savedItems.forEach((item) => {
    const wrapper = document.createElement("article");
    wrapper.className = "saved-item";
    wrapper.innerHTML = `
      <div>
        <p class="saved-item-name">${item.name}</p>
        <div class="saved-item-meta">${formatWeight(item.weightOz)} | ${formatThickness(item.thicknessIn)}</div>
      </div>
      <div class="saved-item-actions">
        <button class="icon-btn" type="button" data-edit-id="${item.id}">Edit</button>
        <button class="icon-btn" type="button" data-delete-id="${item.id}">Delete</button>
      </div>
    `;
    savedItemsRoot.append(wrapper);
  });

  refreshSavedItemSelects();
}

function syncElementLabels() {
  [...elementsRoot.querySelectorAll(".element-card")].forEach((card, index) => {
    card.querySelector(".element-index").textContent = `Element ${index + 1}`;
    card.querySelector(".remove-btn").disabled = state.elements.length === 1;
  });
}

function bindElementCard(card, element) {
  const modeInput = card.querySelector(".input-mode");
  const quantityInput = card.querySelector(".quantity-input");
  const paperFields = card.querySelector(".paper-fields");
  const savedFields = card.querySelector(".saved-fields");
  const paperTypeInput = card.querySelector(".paper-type-input");
  const paperFinishInput = card.querySelector(".paper-finish-input");
  const paperWeightInput = card.querySelector(".paper-weight-input");
  const widthInput = card.querySelector(".width-input");
  const lengthInput = card.querySelector(".length-input");
  const foldTypeInput = card.querySelector(".fold-type-input");
  const savedItemInput = card.querySelector(".saved-item-input");
  const weightOutput = card.querySelector(".element-weight");
  const thicknessOutput = card.querySelector(".element-thickness");
  const elementSummary = card.querySelector(".element-summary");
  const elementSummaryLine = card.querySelector(".element-summary-line");
  const elementSummaryMetrics = card.querySelector(".element-summary-metrics");
  const elementDetails = card.querySelector(".element-details");
  const calcBreakdown = card.querySelector(".calc-breakdown");
  const basisSizeOutput = card.querySelector(".basis-size-output");
  const basisAreaOutput = card.querySelector(".basis-area-output");
  const sqInWeightOutput = card.querySelector(".sq-inch-weight-output");
  const singleSheetOutput = card.querySelector(".single-sheet-output");
  const gsmOutput = card.querySelector(".gsm-output");
  const sheetAreaOutput = card.querySelector(".sheet-area-output");
  const toggleDetailsBtn = card.querySelector(".toggle-details-btn");
  const removeBtn = card.querySelector(".remove-btn");

  modeInput.value = element.mode;
  quantityInput.value = element.quantity;
  paperTypeInput.value = element.paperType;
  paperFinishInput.value = element.paperFinish;
  foldTypeInput.value = element.foldType;
  widthInput.value = element.width;
  lengthInput.value = element.length;
  syncPaperStockOptions(paperTypeInput, paperFinishInput, paperWeightInput, element);

  refreshSavedItemSelects();
  savedItemInput.value = element.savedItemId;
  syncSelectPlaceholderState(modeInput);
  syncSelectPlaceholderState(paperTypeInput);
  syncSelectPlaceholderState(paperFinishInput);
  syncSelectPlaceholderState(paperWeightInput);
  syncSelectPlaceholderState(savedItemInput);

  const renderMode = () => {
    const isSaved = modeInput.value === "saved";
    paperFields.classList.toggle("hidden", isSaved);
    savedFields.classList.toggle("hidden", !isSaved);
    calcBreakdown.classList.toggle("hidden", isSaved);
    paperTypeInput.disabled = isSaved;
    paperFinishInput.disabled = isSaved || paperFinishInput.options.length <= 1;
    paperWeightInput.disabled = isSaved || paperWeightInput.options.length <= 1;
    foldTypeInput.disabled = isSaved;
    widthInput.disabled = isSaved;
    lengthInput.disabled = isSaved;
  };

  const renderCollapsedState = () => {
    elementSummary.classList.toggle("hidden", !element.isCollapsed);
    elementDetails.classList.toggle("hidden", element.isCollapsed);
    toggleDetailsBtn.textContent = element.isCollapsed ? "+" : "-";
    toggleDetailsBtn.setAttribute("aria-label", element.isCollapsed ? "Expand element" : "Collapse element");
  };

  const updateElement = () => {
    element.mode = modeInput.value;
    element.quantity = quantityInput.value;
    element.paperType = paperTypeInput.value;
    element.paperFinish = paperFinishInput.value;
    element.paperWeight = paperWeightInput.value;
    element.foldType = foldTypeInput.value;
    element.width = widthInput.value;
    element.length = lengthInput.value;
    element.savedItemId = savedItemInput.value;

    const result = calculateElement(element);
    element.result = result;
    weightOutput.textContent = formatWeight(result.weightOz);
    thicknessOutput.textContent = formatThickness(result.thicknessIn);
    const summary = getElementSummary(element);
    elementSummaryLine.textContent = summary.title;
    elementSummaryMetrics.textContent = summary.metrics;

    if (element.mode === "paper" && result.reamAreaInSq) {
      basisSizeOutput.textContent = result.basisSizeLabel;
      basisAreaOutput.textContent = formatSquareInches(result.reamAreaInSq);
      sqInWeightOutput.textContent = result.poundsPerSquareIn ? formatWeightPerSquareIn(result.poundsPerSquareIn * 16) : "Point stock";
      singleSheetOutput.textContent = result.poundsPerSheet ? formatWeight(result.poundsPerSheet * 16) : "Weight not derived from pt alone";
      gsmOutput.textContent = result.gsm ? formatGsm(result.gsm) : "N/A";
      sheetAreaOutput.textContent = formatSquareInches(result.sheetAreaInSq);
    } else {
      basisSizeOutput.textContent = "-";
      basisAreaOutput.textContent = "-";
      sqInWeightOutput.textContent = "-";
      singleSheetOutput.textContent = "-";
      gsmOutput.textContent = "-";
      sheetAreaOutput.textContent = "-";
    }

    updateTotals();
  };

  const syncStockSelection = () => {
    syncPaperStockOptions(paperTypeInput, paperFinishInput, paperWeightInput, element);
    syncSelectPlaceholderState(paperTypeInput);
    syncSelectPlaceholderState(paperFinishInput);
    syncSelectPlaceholderState(paperWeightInput);
    renderMode();
    updateElement();
  };

  paperTypeInput.addEventListener("change", syncStockSelection);
  paperFinishInput.addEventListener("change", syncStockSelection);
  paperWeightInput.addEventListener("change", syncStockSelection);

  toggleDetailsBtn.addEventListener("click", () => {
    element.isCollapsed = !element.isCollapsed;
    renderCollapsedState();
  });

  renderMode();
  renderCollapsedState();
  updateElement();

  [modeInput, quantityInput, foldTypeInput, widthInput, lengthInput, savedItemInput]
    .forEach((input) => {
      input.addEventListener("input", () => {
        if (input instanceof HTMLSelectElement) {
          syncSelectPlaceholderState(input);
        }
        renderMode();
        updateElement();
      });
      input.addEventListener("keyup", () => {
        renderMode();
        updateElement();
      });
      input.addEventListener("change", () => {
        if (input instanceof HTMLSelectElement) {
          syncSelectPlaceholderState(input);
        }
        renderMode();
        updateElement();
      });
    });

  removeBtn.addEventListener("click", () => {
    if (state.elements.length === 1) {
      return;
    }

    state.elements = state.elements.filter((item) => item.id !== element.id);
    card.remove();
    syncElementLabels();
    updateTotals();
  });
}

function addElement() {
  if (state.elements.length >= MAX_ELEMENTS) {
    return;
  }

  const element = createElement();
  state.elements.push(element);

  const fragment = elementTemplate.content.cloneNode(true);
  const card = fragment.querySelector(".element-card");
  elementsRoot.append(fragment);

  const mountedCard = elementsRoot.lastElementChild;
  bindElementCard(mountedCard, element);
  syncElementLabels();
  updateTotals();
}

function resetBuild() {
  state.elements = [];
  elementsRoot.innerHTML = "";
  addElement();
}

function initialize() {
  applyDisplaySettings();
  renderSavedItems();
  addElement();

  addElementBtn.addEventListener("click", addElement);
  clearBuildBtn.addEventListener("click", resetBuild);
  settingsBtn.addEventListener("click", openSettingsManager);
  closeSettingsManagerBtn.addEventListener("click", closeSettingsManager);
  manageSavedBtn.addEventListener("click", openSavedManager);
  closeSavedManagerBtn.addEventListener("click", closeSavedManager);
  cancelEditBtn.addEventListener("click", clearSavedItemForm);
  settingsManager.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (target.dataset.closeSettingsManager) {
      closeSettingsManager();
    }
  });
  savedManager.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (target.dataset.closeSavedManager) {
      closeSavedManager();
    }
  });

  settingsForm.addEventListener("submit", (event) => {
    event.preventDefault();

    state.settings = {
      weightUnit: weightUnitInput.value,
      thicknessUnit: thicknessUnitInput.value,
      fudgeFactor: Number(fudgeFactorInput.value) > 0 ? Number(fudgeFactorInput.value) : DEFAULT_SETTINGS.fudgeFactor
    };

    persistSettings();
    applyDisplaySettings();
    closeSettingsManager();
  });

  customItemForm.addEventListener("submit", (event) => {
    event.preventDefault();

    const formData = new FormData(customItemForm);
    const name = String(formData.get("name") ?? "").trim();
    const weightOz = convertWeightToOz(Number(formData.get("weight")));
    const thicknessIn = convertThicknessToIn(Number(formData.get("thickness")));

    if (!name || weightOz < 0 || thicknessIn < 0) {
      return;
    }

    if (editingSavedItemId) {
      state.savedItems = state.savedItems.map((item) => item.id === editingSavedItemId
        ? { ...item, name, weightOz, thicknessIn }
        : item);
    } else {
      state.savedItems.unshift({
        id: createId(),
        name,
        weightOz,
        thicknessIn
      });
    }

    persistSavedItems();
    renderSavedItems();

    state.elements.forEach((element, index) => {
      if (!element.savedItemId && index === 0) {
        element.savedItemId = state.savedItems[0].id;
      }
    });

    document.querySelectorAll(".saved-item-input").forEach((select) => {
      if (!select.value) {
        select.value = state.savedItems[0].id;
      }
      select.dispatchEvent(new Event("change"));
    });

    clearSavedItemForm();
  });

  savedItemsRoot.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const editId = target.dataset.editId;
    if (editId) {
      const item = state.savedItems.find((savedItem) => savedItem.id === editId);
      if (item) {
        startEditingSavedItem(item);
      }
      return;
    }

    const deleteId = target.dataset.deleteId;
    if (!deleteId) {
      return;
    }

    const itemToDelete = state.savedItems.find((item) => item.id === deleteId);
    if (!itemToDelete) {
      return;
    }

    const confirmed = window.confirm(`Are you sure you want to delete "${itemToDelete.name}"?`);
    if (!confirmed) {
      return;
    }

    state.savedItems = state.savedItems.filter((item) => item.id !== deleteId);
    if (state.savedItems.length === 0) {
      state.savedItems = [...defaultSavedItems];
    }

    state.elements.forEach((element) => {
      if (element.savedItemId === deleteId) {
        element.savedItemId = state.savedItems[0]?.id ?? "";
      }
    });

    persistSavedItems();
    renderSavedItems();
    document.querySelectorAll(".saved-item-input").forEach((select, index) => {
      select.value = state.elements[index]?.savedItemId ?? state.savedItems[0]?.id ?? "";
      select.dispatchEvent(new Event("change"));
    });
  });
}

initialize();
