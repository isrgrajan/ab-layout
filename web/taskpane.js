/*
File: taskpane.js
Path: /web/taskpane.js
Version: 3.1.0 (FINAL)

Description:
Production-grade layout engine with ribbon integration, validation, and UI state control

Maintainer:
RatioJuris
*/

/* ===== GLOBAL STATE ===== */

let isOfficeReady = false;
let dataStore = {};
let currentLayout = null;
let previousLayout = null;

const LAYOUTS_URL =
  "https://ratiojuris.github.io/ab-layout/layouts/layouts.json?v=3.1.0";

/* ===== INIT ===== */

Office.onReady(async (info) => {

  if (info.host === Office.HostType.Word) {
    isOfficeReady = true;

    registerRibbonActions();   // 🔥 critical
    init();

  } else {
    uiSetStatus("Open inside Microsoft Word", "error");
  }

});

function init() {
  uiSetEngineState("Loading...");
  loadLayouts();
}

/* ===== UI SAFE WRAPPER ===== */

function setStatus(message, type = "info") {
  if (typeof uiSetStatus === "function") {
    uiSetStatus(message, type);
  } else {
    const el = document.getElementById("status");
    if (el) el.innerText = message;
  }
}

/* ===== LOAD LAYOUTS ===== */

async function loadLayouts() {
  try {
    setStatus("Loading layouts...", "info");

    const res = await fetch(LAYOUTS_URL);
    if (!res.ok) throw new Error("Fetch failed");

    dataStore = await res.json();

    validateSchema(dataStore);

    populateLayouts();

    uiEnableApply(true);
    uiSetEngineState("Ready");

    setStatus("Ready", "success");

  } catch (err) {
    console.error(err);
    setStatus("Failed to load layouts", "error");
    uiSetEngineState("Error");
  }
}

/* ===== VALIDATION ===== */

function validateSchema(store) {
  if (!store.layouts || !store.presets) {
    throw new Error("Invalid schema");
  }
}

function validateLayout(layout) {

  if (!layout) return fail("Layout missing");

  if (!isNumber(layout.width) || !isNumber(layout.height)) {
    return fail("Invalid page size");
  }

  const checkMargins = (m) =>
    m &&
    ["top", "bottom", "left", "right"].every(
      k => isNumber(m[k]) && m[k] >= 0
    );

  if (layout.multiPage) {
    if (!checkMargins(layout.firstPage) || !checkMargins(layout.otherPages)) {
      return fail("Invalid multi-page margins");
    }
  } else {
    if (!checkMargins(layout.margins)) {
      return fail("Invalid margins");
    }
  }

  return { valid: true };
}

function isNumber(v) {
  return typeof v === "number" && !isNaN(v);
}

function fail(msg) {
  return { valid: false, error: msg };
}

/* ===== UI ===== */

function populateLayouts() {
  const select = document.getElementById("courtSelect");
  if (!select) return;

  select.innerHTML = "";

  dataStore.layouts.forEach((layout, i) => {
    const option = document.createElement("option");
    option.value = i;
    option.textContent = layout.name;
    select.appendChild(option);
  });
}

/* ===== RESOLUTION ===== */

function getSelectedLayout() {
  const select = document.getElementById("courtSelect");
  if (!select) return null;

  const index = select.value;

  const item = dataStore.layouts[index];
  if (!item) return null;

  const preset = dataStore.presets[item.preset];
  if (!preset) return null;

  return {
    id: item.id,
    name: item.name,
    ...preset
  };
}

/* ===== APPLY CORE ===== */

async function applySelected() {

  if (!isOfficeReady) {
    setStatus("Open inside Microsoft Word", "error");
    return;
  }

  const layout = getSelectedLayout();
  if (!layout) return;

  if (currentLayout?.id === layout.id) {
    setStatus("Already applied", "info");
    return;
  }

  const check = validateLayout(layout);
  if (!check.valid) {
    setStatus(check.error, "error");
    return;
  }

  uiSetEngineState("Applying...");
  setStatus("Applying layout...", "info");

  try {

    await Word.run(async (context) => {

      const sections = context.document.sections;
      sections.load("items/pageSetup");

      await context.sync();

      if (!sections.items.length) return;

      const first = sections.items[0].pageSetup;

      previousLayout = {
        width: first.pageWidth,
        height: first.pageHeight,
        top: first.topMargin,
        bottom: first.bottomMargin,
        left: first.leftMargin,
        right: first.rightMargin
      };

      sections.items.forEach((section, i) => {

        let margins = layout.margins;

        if (layout.multiPage) {
          margins = i === 0 ? layout.firstPage : layout.otherPages;
        }

        if (!margins) return;

        section.pageSetup.pageWidth = layout.width;
        section.pageSetup.pageHeight = layout.height;

        section.pageSetup.topMargin = margins.top;
        section.pageSetup.bottomMargin = margins.bottom;
        section.pageSetup.leftMargin = margins.left;
        section.pageSetup.rightMargin = margins.right;

      });

      await context.sync();
    });

    currentLayout = layout;

    uiEnableUndo(true);
    uiSetEngineState("Active");

    setStatus("Applied: " + layout.name, "success");

  } catch (err) {

    console.error(err);

    setStatus("Failed to apply layout", "error");
    uiSetEngineState("Error");

    if (previousLayout) undoLayout();
  }
}

/* ===== APPLY BY ID (RIBBON) ===== */

function applyById(id) {

  if (!dataStore.layouts) return;

  const index = dataStore.layouts.findIndex(l => l.id === id);
  if (index === -1) return;

  const select = document.getElementById("courtSelect");
  if (select) select.value = index;

  applySelected();
}

/* ===== UNDO ===== */

async function undoLayout() {

  if (!previousLayout) {
    setStatus("No previous layout", "info");
    return;
  }

  uiSetEngineState("Reverting...");

  try {

    await Word.run(async (context) => {

      const sections = context.document.sections;
      sections.load("items");

      await context.sync();

      sections.items.forEach((section) => {

        section.pageSetup.pageWidth = previousLayout.width;
        section.pageSetup.pageHeight = previousLayout.height;

        section.pageSetup.topMargin = previousLayout.top;
        section.pageSetup.bottomMargin = previousLayout.bottom;
        section.pageSetup.leftMargin = previousLayout.left;
        section.pageSetup.rightMargin = previousLayout.right;

      });

      await context.sync();
    });

    currentLayout = null;

    uiEnableUndo(false);
    uiSetEngineState("Ready");

    setStatus("Layout restored", "success");

  } catch (err) {
    console.error(err);
    setStatus("Undo failed", "error");
  }
}

/* ===== RIBBON ACTION REGISTRATION ===== */

function registerRibbonActions() {

  if (!Office.actions) return;

  Office.actions.associate("applyHighCourt", () => applyById("hc_standard"));
  Office.actions.associate("applyDistrictCourt", () => applyById("district_standard"));
  Office.actions.associate("applyAffidavit", () => applyById("affidavit"));
  Office.actions.associate("undoLayout", undoLayout);
}
