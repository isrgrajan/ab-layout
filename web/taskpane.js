/*
AB Layout (Advocate Benefit Layout)
Version: 2.0.0

Description:
Validated, points-based layout engine for Microsoft Word Add-in.

Maintainer:
RatioJuris
*/

/* ===== STATE ===== */

let isOfficeReady = false;
let dataStore = {};
let currentLayout = null;
let previousLayout = null;

const LAYOUTS_URL =
  "https://ratiojuris.github.io/ab-layout/layouts/layouts.json?v=2.0.0";

/* ===== INIT ===== */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    isOfficeReady = true;
    init();
  } else {
    setStatus("Open inside Microsoft Word");
  }
});

function init() {
  loadLayouts();
}

/* ===== UI ===== */

function setStatus(message) {
  const el = document.getElementById("status");
  if (el) el.innerText = message;
}

/* ===== DATA LOADING ===== */

async function loadLayouts() {
  try {
    setStatus("Loading layouts...");

    const res = await fetch(LAYOUTS_URL);
    if (!res.ok) throw new Error("Failed to fetch layouts");

    dataStore = await res.json();

    validateSchema(dataStore);

    populateLayouts();

    document.getElementById("applyBtn").disabled = false;

    setStatus("Ready");
  } catch (err) {
    console.error(err);
    setStatus("Failed to load layouts");
  }
}

/* ===== SCHEMA VALIDATION ===== */

function validateSchema(store) {
  if (!store.layouts || !store.presets) {
    throw new Error("Invalid layout schema");
  }
}

/* ===== UI POPULATION ===== */

function populateLayouts() {
  const select = document.getElementById("courtSelect");
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
  const index = document.getElementById("courtSelect").value;

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

/* ===== VALIDATION ENGINE ===== */

function validateLayout(layout) {
  if (!layout) return fail("Layout missing");

  if (!isNumber(layout.width) || !isNumber(layout.height)) {
    return fail("Invalid page size");
  }

  if (layout.width <= 0 || layout.height <= 0) {
    return fail("Page size must be positive");
  }

  const checkMargins = (m) => {
    if (!m) return false;

    return ["top", "bottom", "left", "right"].every(
      k => isNumber(m[k]) && m[k] >= 0
    );
  };

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

/* ===== APPLY ENGINE ===== */

async function applySelected() {
  if (!isOfficeReady) {
    setStatus("Open inside Microsoft Word");
    return;
  }

  const layout = getSelectedLayout();
  if (!layout) return;

  if (currentLayout?.id === layout.id) {
    setStatus("Already applied");
    return;
  }

  const check = validateLayout(layout);

  if (!check.valid) {
    setStatus("Error: " + check.error);
    console.error(check.error);
    return;
  }

  try {
    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items/pageSetup");

      await context.sync();

      if (!sections.items.length) return;

      /* SAVE STATE */
      const first = sections.items[0].pageSetup;

      previousLayout = {
        width: first.pageWidth,
        height: first.pageHeight,
        top: first.topMargin,
        bottom: first.bottomMargin,
        left: first.leftMargin,
        right: first.rightMargin
      };

      /* APPLY */
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
    setStatus("Applied: " + layout.name);

  } catch (err) {
    console.error(err);
    setStatus("Failed to apply layout");

    if (previousLayout) undoLayout();
  }
}

/* ===== UNDO ENGINE ===== */

async function undoLayout() {
  if (!previousLayout) {
    setStatus("No previous layout");
    return;
  }

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
    setStatus("Layout restored");

  } catch (err) {
    console.error(err);
    setStatus("Undo failed");
  }
}
