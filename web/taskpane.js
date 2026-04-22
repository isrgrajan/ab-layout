/*
AB Layout (Advocate Benefit Layout)
Version: 1.0.0

Description:
Production-ready layout engine for Microsoft Word Add-in.

Supports:
- Preset-based layouts
- Unit conversion (inch/cm)
- Multi-page layouts (affidavit-ready)
- Safe apply + undo
- Clean, extensible architecture

Author:
Bee Isrg Rajan

Repository:
https://github.com/isrgrajan/ab-layout
*/

Office.onReady(() => {
  loadLayouts();
});

const LAYOUTS_URL =
  "https://isrgrajan.github.io/ab-layout/layouts/layouts.json?v=1.0.0";

let dataStore = {};
let currentLayout = null;
let previousLayout = null;

/* ===== Utils ===== */

function setStatus(message) {
  const el = document.getElementById("status");
  if (el) el.innerText = message;
}

function toPoints(value, unit = "inch") {
  if (unit === "cm") return value * 28.3465;
  return value * 72;
}

/* ===== Load Layouts ===== */

async function loadLayouts() {
  try {
    setStatus("Loading layouts...");

    const response = await fetch(LAYOUTS_URL);
    if (!response.ok) throw new Error("Failed to fetch layouts");

    dataStore = await response.json();

    populateStates();

    document.getElementById("applyBtn").disabled = false;

    setStatus("Ready");
  } catch (error) {
    console.error(error);
    setStatus("Failed to load layouts");
  }
}

/* ===== Populate UI ===== */

function populateStates() {
  const stateSelect = document.getElementById("stateSelect");
  stateSelect.innerHTML = "";

  (dataStore.states || []).forEach((state, index) => {
    const option = document.createElement("option");
    option.value = index;
    option.textContent = state.name;
    stateSelect.appendChild(option);
  });

  stateSelect.addEventListener("change", populateCourts);

  populateCourts();
}

function populateCourts() {
  const stateIndex = document.getElementById("stateSelect").value;
  const courtSelect = document.getElementById("courtSelect");

  courtSelect.innerHTML = "";

  const courts = dataStore.states[stateIndex]?.courts || [];

  courts.forEach((court, index) => {
    const option = document.createElement("option");
    option.value = index;
    option.textContent = court.name;
    courtSelect.appendChild(option);
  });
}

/* ===== Layout Resolution ===== */

function getSelectedLayout() {
  const stateIndex = document.getElementById("stateSelect").value;
  const courtIndex = document.getElementById("courtSelect").value;

  const court = dataStore.states[stateIndex]?.courts[courtIndex];
  if (!court) return null;

  const preset = dataStore.presets?.[court.preset];
  if (!preset) return null;

  return {
    id: court.preset,
    name: court.name,
    ...preset
  };
}

/* ===== Apply Layout ===== */

async function applySelected() {
  const layout = getSelectedLayout();
  if (!layout) return;

  if (currentLayout && currentLayout.id === layout.id) {
    setStatus("Already applied");
    return;
  }

  try {
    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items/pageSetup");

      await context.sync();

      if (!sections.items.length) return;

      /* Save previous layout */
      const first = sections.items[0].pageSetup;

      previousLayout = {
        width: first.pageWidth,
        height: first.pageHeight,
        top: first.topMargin,
        bottom: first.bottomMargin,
        left: first.leftMargin,
        right: first.rightMargin
      };

      const unit = layout.unit || dataStore._meta?.unit || "inch";

      sections.items.forEach((section, index) => {
        let margins = layout.margins;

        /* Multi-page handling */
        if (layout.multiPage) {
          margins = index === 0 ? layout.firstPage : layout.otherPages;
        }

        section.pageSetup.pageWidth = toPoints(layout.width, unit);
        section.pageSetup.pageHeight = toPoints(layout.height, unit);

        section.pageSetup.topMargin = toPoints(margins.top, unit);
        section.pageSetup.bottomMargin = toPoints(margins.bottom, unit);
        section.pageSetup.leftMargin = toPoints(margins.left, unit);
        section.pageSetup.rightMargin = toPoints(margins.right, unit);
      });

      await context.sync();
    });

    currentLayout = layout;
    setStatus("Applied: " + layout.name);
  } catch (error) {
    console.error(error);
    setStatus("Failed to apply layout");
  }
}

/* ===== Undo Layout ===== */

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
  } catch (error) {
    console.error(error);
    setStatus("Undo failed");
  }
}
