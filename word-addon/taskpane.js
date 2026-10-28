// Ghost AI - Word Add-in Controller (Stealth Edition)
const STORAGE_KEY = "ghost_gemini_api_key";
const STORAGE_THEME = "ghost_addon_theme";
const STORAGE_OPACITY = "ghost_addon_opacity";

let currentRawMarkdown = "";
let currentRenderedHtml = "";
let currentPertOoxml = null;
let chatHistory = [];
let chatTurns = [];
let activeTurnIndex = -1;

// Elements
const stealthContent = document.getElementById("stealth-content");
const panicZone = document.getElementById("panic-zone");
const promptInput = document.getElementById("prompt-input");
const exerciseType = document.getElementById("exercise-type");
const btnSolve = document.getElementById("btn-solve");
const btnClear = document.getElementById("btn-clear");
const btnGetSelection = document.getElementById("btn-get-selection");
const btnInsertWord = document.getElementById("btn-insert-word");
const btnInsertDrawing = document.getElementById("btn-insert-drawing");
const btnCopyClipboard = document.getElementById("btn-copy-clipboard");
const resultContainer = document.getElementById("result-container");
const resultView = document.getElementById("result-view");
const loadingBox = document.getElementById("loading");

// History Navigation & Pinned Header Elements (v3.7)
const pinnedHistoryBanner = document.getElementById("pinned-history-banner");
const btnPinnedPrev = document.getElementById("btn-pinned-prev");
const pinnedPrevTitle = document.getElementById("pinned-prev-title");
const pinnedNextBanner = document.getElementById("pinned-next-banner");
const btnPinnedNext = document.getElementById("btn-pinned-next");
const pinnedNextTitle = document.getElementById("pinned-next-title");
const btnTurnPrev = document.getElementById("btn-turn-prev");
const btnTurnNext = document.getElementById("btn-turn-next");
const turnBadge = document.getElementById("turn-badge");
const turnUserBox = document.getElementById("turn-user-box");
const turnUserText = document.getElementById("turn-user-text");

// Stealth Bar Controls
const stealthBar = document.getElementById("stealth-bar");
const btnTheme = document.getElementById("btn-theme");
const opacitySlider = document.getElementById("opacity-slider");
const opacityVal = document.getElementById("opacity-val");

// Force top bar to strictly 0% opacity by default (reveals only on hover)
if (stealthBar) {
  stealthBar.style.opacity = "0";
  stealthBar.style.transition = "opacity 0.15s ease";

  stealthBar.addEventListener("mouseenter", () => {
    stealthBar.style.opacity = "1";
  });
  stealthBar.addEventListener("mouseleave", () => {
    stealthBar.style.opacity = "0";
  });
}

// Settings Modal elements
const btnSettings = document.getElementById("btn-settings");
const apiModal = document.getElementById("api-modal");
const modalKeyInput = document.getElementById("modal-key-input");
const btnModalSave = document.getElementById("btn-modal-save");
const btnModalCancel = document.getElementById("btn-modal-cancel");

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Add-in ready inside Word.");
  }
});

// Theme Management (Monochrome Moon/Sun symbol)
function applyTheme(theme) {
  if (theme === "dark") {
    document.body.classList.remove("theme-light");
    document.body.classList.add("theme-dark");
    btnTheme.textContent = "☼";
    btnTheme.title = "Cambiar a modo claro (☼)";
  } else {
    document.body.classList.remove("theme-dark");
    document.body.classList.add("theme-light");
    btnTheme.textContent = "☾";
    btnTheme.title = "Cambiar a modo oscuro (☾)";
  }
  localStorage.setItem(STORAGE_THEME, theme);
}

btnTheme.addEventListener("click", () => {
  const isDark = document.body.classList.contains("theme-dark");
  applyTheme(isDark ? "light" : "dark");
});

// Opacity & Panic Mode (Esc & Right-Edge Proximity)
let lastActiveOpacity = 20;
let isPanicHidden = false;

function applyOpacity(val, remember = true) {
  const clamped = Math.max(0, Math.min(100, parseInt(val, 10) || 0));
  if (stealthContent) {
    stealthContent.style.opacity = (clamped / 100).toString();
    stealthContent.style.pointerEvents = clamped === 0 ? "none" : "auto";
  }
  opacityVal.textContent = clamped + "%";
  opacitySlider.value = clamped;

  if (clamped === 0) {
    isPanicHidden = true;
    document.body.classList.add("panic-mode");
  } else {
    isPanicHidden = false;
    document.body.classList.remove("panic-mode");
    if (remember) {
      lastActiveOpacity = clamped;
      localStorage.setItem(STORAGE_OPACITY, clamped);
    }
  }
}

function triggerPanic() {
  if (!isPanicHidden && (parseInt(opacitySlider.value, 10) > 0)) {
    lastActiveOpacity = Math.max(20, parseInt(opacitySlider.value, 10) || 20);
    applyOpacity(0, false);
  }
}

function restorePanic() {
  const restoreVal = Math.max(20, lastActiveOpacity || 20);
  applyOpacity(restoreVal, true);
}

function togglePanic() {
  if (isPanicHidden || parseInt(opacitySlider.value, 10) === 0) {
    restorePanic();
  } else {
    triggerPanic();
  }
}

// Right-Edge Proximity Trigger: moving mouse to extreme right edge drops to 0% and locks
if (panicZone) {
  panicZone.addEventListener("mouseenter", () => {
    if (!isPanicHidden) triggerPanic();
  });
}

window.addEventListener("mousemove", (e) => {
  // Only trigger panic if cursor is below the top bar (clientY > 45) and at the right edge
  if (!isPanicHidden && e.clientY > 45 && e.clientX >= (window.innerWidth - 18)) {
    triggerPanic();
  }
});

// Global key listener inside add-in for 'Escape' key
window.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    e.preventDefault();
    togglePanic();
  }
});

opacitySlider.addEventListener("input", (e) => {
  applyOpacity(e.target.value, true);
});

// Load saved preferences
const savedTheme = localStorage.getItem(STORAGE_THEME) || "light";
applyTheme(savedTheme);

const savedOpacity = localStorage.getItem(STORAGE_OPACITY) || "20";
lastActiveOpacity = Math.max(20, parseInt(savedOpacity, 10) || 20);
applyOpacity(lastActiveOpacity, true);

// API Key Management
function getApiKey() {
  return localStorage.getItem(STORAGE_KEY) || "";
}

btnSettings.addEventListener("click", () => {
  modalKeyInput.value = getApiKey();
  apiModal.classList.add("active");
  modalKeyInput.focus();
});

btnModalCancel.addEventListener("click", () => {
  apiModal.classList.remove("active");
});

btnModalSave.addEventListener("click", () => {
  const val = modalKeyInput.value.trim();
  if (!val) {
    alert("Por favor ingresa una API Key válida de Gemini.");
    return;
  }
  localStorage.setItem(STORAGE_KEY, val);
  apiModal.classList.remove("active");
  alert("API Key guardada con éxito.");
});

// Clear input and reset conversational memory
btnClear.addEventListener("click", () => {
  promptInput.value = "";
  resultContainer.style.display = "none";
  currentRawMarkdown = "";
  currentRenderedHtml = "";
  currentPertOoxml = null;
  chatHistory = [];
  chatTurns = [];
  activeTurnIndex = -1;
  btnSolve.textContent = "Resolver Paso a Paso";
  promptInput.placeholder = "Selecciona el ejercicio en tu hoja de Word y toca 'Tomar de Word', o pega la consigna aquí...";
  if (btnInsertDrawing) btnInsertDrawing.style.display = "none";
  if (pinnedHistoryBanner) pinnedHistoryBanner.style.display = "none";
  if (pinnedNextBanner) pinnedNextBanner.style.display = "none";
  if (resultView) resultView.innerHTML = "";
});

// Take selection from Word document
btnGetSelection.addEventListener("click", async () => {
  if (typeof Word !== "undefined" && Word.run) {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        if (selection.text && selection.text.trim()) {
          promptInput.value = selection.text.trim();
        } else {
          alert("Resalta primero el texto de la consigna en tu hoja de Word.");
        }
      });
      return;
    } catch (e) {
      console.warn("Word.run getSelection fallback:", e);
    }
  }

  // Fallback Office.js API
  if (typeof Office !== "undefined" && Office.context && Office.context.document) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        promptInput.value = result.value.trim();
      } else {
        alert("Resalta primero el texto de la consigna en tu hoja de Word.");
      }
    });
  } else {
    alert("Para usar esta función, abre el complemento dentro de Word.");
  }
});

// Helper function to prepare native Word HTML elements (Tables, Boxes, Arrows)
function prepareWordHtml(html) {
  let clean = html;

  // 0. Remove hidden HTML comments (e.g. PERT_DATA)
  clean = clean.replace(/<!--[\s\S]*?-->/g, "");

  // 1. Remove code blocks (<pre><code>) if any, converting them into clean normal text
  clean = clean.replace(/<pre><code[^>]*>([\s\S]*?)<\/code><\/pre>/gi, (match, code) => {
    return `<div style="font-family: Calibri, sans-serif; padding: 8px 0; margin: 8px 0; line-height: 1.4;">${code}</div>`;
  });

  // 2. Remove LaTeX commands, raw math codes and dollar signs
  clean = clean.replace(/\\implies/g, " ➔ ")
               .replace(/\\rightarrow/g, " ➔ ")
               .replace(/-->/g, " ➔ ")
               .replace(/->/g, " ➔ ")
               .replace(/\\max/g, "max")
               .replace(/\\min/g, "min")
               .replace(/\\text\{([^}]+)\}/g, "$1")
               .replace(/\$([^\$]+)\$/g, "$1");

  // 3. Format square brackets like [ A: 3 meses ] into real bordered Word boxes (Cuadros de Word)
  clean = clean.replace(/\[\s*([^\]\<\>]+?)\s*\]/g, (match, text) => {
    if (text.startsWith("http") || text.length > 60 || text.includes("\n")) return match;
    const isCritical = text.includes("*") || text.toLowerCase().includes("crític");
    const bg = isCritical ? "#e2e8f0" : "#ffffff";
    const border = isCritical ? "2px solid #000000" : "1.5px solid #000000";
    return `<span style="display: inline-block; border: ${border}; background-color: ${bg}; padding: 3px 8px; margin: 3px 4px; font-weight: bold; font-family: Calibri, sans-serif; font-size: 10pt; text-align: center; border-radius: 2px;">${text}</span>`;
  });

  // 4. Ensure real native Word tables with clean black borders and cell shading
  clean = clean.replace(/<table[^>]*>/gi, '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1.5px solid #000000; margin: 12px 0; font-family: Calibri, sans-serif; font-size: 10pt;">');
  clean = clean.replace(/<th[^>]*>/gi, '<th style="border: 1px solid #000000; background-color: #f2f2f2; font-weight: bold; padding: 6px 8px; text-align: center; color: #000000;">');
  clean = clean.replace(/<td[^>]*>/gi, '<td style="border: 1px solid #000000; padding: 5px 8px; text-align: center; color: #000000;">');

  return clean;
}

// Insert solution directly into Word document
btnInsertWord.addEventListener("click", async () => {
  if (!currentRenderedHtml) {
    alert("No hay desarrollo generado para insertar.");
    return;
  }

  // Academic styling for native Word document insertion
  const styledHtml = `
    <div style="font-family: Calibri, 'Segoe UI', sans-serif; font-size: 11pt; color: #000000; line-height: 1.35;">
      <h3 style="color: #000000; border-bottom: 2px solid #000000; padding-bottom: 4px; margin-top: 14px;">
        Desarrollo y Resolución Técnica
      </h3>
      ${currentRenderedHtml}
    </div>
  `;

  if (typeof Word !== "undefined" && Word.run) {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertHtml(styledHtml, Word.InsertLocation.replace);
        await context.sync();
      });
      alert("Desarrollo y tablas insertados exitosamente en tu documento de Word.");
      return;
    } catch (e) {
      console.warn("Word.run insertHtml fallback:", e);
    }
  }

  // Fallback Office.js API
  if (typeof Office !== "undefined" && Office.context && Office.context.document) {
    Office.context.document.setSelectedDataAsync(styledHtml, { coercionType: Office.CoercionType.Html }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        alert("Desarrollo insertado exitosamente en tu Word.");
      } else {
        alert("Error al insertar en Word: " + result.error.message);
      }
    });
  } else {
    // If testing in standard browser outside Word
    navigator.clipboard.writeText(currentRawMarkdown);
    alert("Copiado al portapapeles. Pégalo en tu Word con Ctrl+V.");
  }
});

// Copy to clipboard
btnCopyClipboard.addEventListener("click", () => {
  if (!currentRawMarkdown) return;
  navigator.clipboard.writeText(currentRawMarkdown).then(() => {
    alert("Desarrollo copiado al portapapeles.");
  });
});

// ==========================================
// Full PERT Graph & DrawingML OOXML Generator
// ==========================================
function parsePertTable(markdownText) {
  // 1. Try parsing from hidden HTML comment first
  const commentMatch = markdownText.match(/<!--\s*PERT_DATA:\s*([\s\S]*?)-->/);
  if (commentMatch) {
    try {
      const data = JSON.parse(commentMatch[1].trim());
      if (Array.isArray(data) && data.length > 0) return data;
    } catch (e) {
      console.warn("Failed to parse PERT_DATA comment:", e);
    }
  }

  // 2. Fallback: Parse markdown table
  const lines = markdownText.split("\n");
  const activities = [];
  let tableFound = false;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed.startsWith("|") || !trimmed.endsWith("|")) {
      if (tableFound && activities.length > 0) break;
      continue;
    }

    const cells = trimmed.split("|").slice(1, -1).map(c => c.trim().replace(/\*\*/g, "").replace(/\*/g, ""));
    if (cells.length < 8) continue;

    const firstCell = cells[0].toLowerCase();
    if (firstCell.includes("actividad") || firstCell.includes("tarea") || firstCell.includes("nodo")) {
      tableFound = true;
      continue;
    }
    if (cells[0].includes("---") || cells[0].includes(":-")) {
      continue;
    }

    if (!tableFound) continue;

    const id = cells[0].replace(/[^a-zA-Z0-9_]/g, "").toUpperCase();
    if (!id || id.length > 10) continue;

    const predStr = cells[1];
    const preds = (predStr === "-" || predStr.toLowerCase().includes("ningun") || !predStr)
      ? []
      : predStr.split(/[,;\s]+/).map(p => p.replace(/[^a-zA-Z0-9_]/g, "").toUpperCase()).filter(Boolean);

    const t = parseFloat(cells[2]) || 0;
    const es = parseFloat(cells[3]) || 0;
    const ef = parseFloat(cells[4]) || 0;
    const ls = parseFloat(cells[5]) || 0;
    const lf = parseFloat(cells[6]) || 0;
    const h = parseFloat(cells[7]) || (lf - ef);
    const crit = (cells[8] && (cells[8].toLowerCase().includes("s") || cells[8].toLowerCase().includes("crít"))) || (h === 0);

    activities.push({ id, act: id, t, pred: preds, es, ef, ls, lf, h, crit });
  }

  return activities;
}

function buildFullPertOoxmlFromActivities(rawActivities) {
  if (!rawActivities || rawActivities.length === 0) return null;

  const acts = rawActivities.map(a => ({ ...a, pred: [...a.pred] }));

  // Ensure INICIO node
  const roots = acts.filter(a => a.pred.length === 0 && a.id !== "INICIO");
  let hasInicio = acts.some(a => a.id === "INICIO");
  if (!hasInicio && roots.length > 0) {
    acts.unshift({
      id: "INICIO",
      act: "INICIO",
      t: 0,
      pred: [],
      es: 0,
      ef: 0,
      ls: 0,
      lf: 0,
      h: 0,
      crit: true
    });
    roots.forEach(r => r.pred.push("INICIO"));
  }

  // Ensure FIN node
  const allPreds = new Set();
  acts.forEach(a => a.pred.forEach(p => allPreds.add(p)));
  const terminals = acts.filter(a => !allPreds.has(a.id) && a.id !== "FIN");
  let hasFin = acts.some(a => a.id === "FIN");
  if (!hasFin && terminals.length > 0) {
    const maxEf = Math.max(...acts.map(a => a.ef || 0));
    acts.push({
      id: "FIN",
      act: "FIN",
      t: 0,
      pred: terminals.map(t => t.id),
      es: maxEf,
      ef: maxEf,
      ls: maxEf,
      lf: maxEf,
      h: 0,
      crit: true
    });
  }

  // 1. Successor mapping
  const succs = {};
  acts.forEach(a => { succs[a.id] = []; });
  acts.forEach(a => {
    a.pred.forEach(p => {
      if (succs[p]) succs[p].push(a.id);
    });
  });

  // 2. Topological Levels (Col X)
  const levels = {};
  function getLevel(actId, visited = new Set()) {
    if (levels[actId] !== undefined) return levels[actId];
    if (visited.has(actId)) return 0;
    visited.add(actId);

    const act = acts.find(a => a.id === actId);
    if (!act || act.pred.length === 0) {
      levels[actId] = 0;
      return 0;
    }

    const lvl = Math.max(...act.pred.map(p => getLevel(p, visited))) + 1;
    levels[actId] = lvl;
    return lvl;
  }
  acts.forEach(a => getLevel(a.id));

  // 3. Find all paths from INICIO to FIN
  const allPaths = [];
  function dfsPaths(curr, path) {
    if (curr === "FIN") {
      allPaths.push([...path]);
      return;
    }
    const children = succs[curr] || [];
    for (const s of children) {
      dfsPaths(s, [...path, s]);
    }
  }
  dfsPaths("INICIO", ["INICIO"]);

  function isPathCrit(path) {
    return path.every(nid => {
      const a = acts.find(x => x.id === nid);
      return a ? a.crit : false;
    });
  }

  // Group paths by root branch
  const rootBranches = [];
  allPaths.forEach(p => {
    const rootNode = p.length > 1 ? p[1] : "FIN";
    if (!rootBranches.includes(rootNode)) {
      rootBranches.push(rootNode);
    }
  });

  const critRoot = allPaths.find(p => isPathCrit(p) && p.length > 1)?.[1] || rootBranches[0];
  const reorderedRootBranches = [critRoot, ...rootBranches.filter(r => r !== critRoot)];

  const sortedPaths = [];
  reorderedRootBranches.forEach(rb => {
    const rbPaths = allPaths.filter(p => p.length > 1 && p[1] === rb);
    rbPaths.sort((p1, p2) => {
      const c1 = isPathCrit(p1) ? 1 : 0;
      const c2 = isPathCrit(p2) ? 1 : 0;
      if (c1 !== c2) return c2 - c1;
      return p2.length - p1.length;
    });
    sortedPaths.push(...rbPaths);
  });

  const nodePathCount = {};
  sortedPaths.forEach(p => {
    p.forEach(nid => {
      nodePathCount[nid] = (nodePathCount[nid] || 0) + 1;
    });
  });

  const tracks = {};
  let currentTrack = 0;

  sortedPaths.forEach(p => {
    const middle = p.filter(nid => nid !== "INICIO" && nid !== "FIN");
    const unassigned = middle.filter(nid => tracks[nid] === undefined && nodePathCount[nid] === 1);
    if (unassigned.length > 0) {
      unassigned.forEach(nid => {
        tracks[nid] = currentTrack;
      });
      currentTrack++;
    }
  });

  const multiNodes = acts.filter(a => tracks[a.id] === undefined && a.id !== "INICIO" && a.id !== "FIN");
  multiNodes.forEach(a => {
    const nid = a.id;
    const chTracks = (succs[nid] || []).filter(c => tracks[c] !== undefined).map(c => tracks[c]);
    const prTracks = (a.pred || []).filter(p => tracks[p] !== undefined).map(p => tracks[p]);
    const combined = [...chTracks, ...prTracks];
    if (combined.length > 0) {
      tracks[nid] = combined.reduce((sum, v) => sum + v, 0) / combined.length;
    } else {
      tracks[nid] = currentTrack++;
    }
  });

  const incTracks = (succs["INICIO"] || []).filter(c => tracks[c] !== undefined).map(c => tracks[c]);
  tracks["INICIO"] = incTracks.length > 0 ? (incTracks.reduce((s, v) => s + v, 0) / incTracks.length) : 0;

  const finAct = acts.find(a => a.id === "FIN");
  const finTracks = (finAct?.pred || []).filter(p => tracks[p] !== undefined).map(p => tracks[p]);
  tracks["FIN"] = finTracks.length > 0 ? (finTracks.reduce((s, v) => s + v, 0) / finTracks.length) : 0;

  // 4. Geometry & Compact Spacing (Fits inside standard Word margins)
  const nodeW = 480000;      // ~1.33 cm
  const boxW = Math.floor(nodeW / 2); // 240000
  const boxH = 140000;      // ~0.39 cm
  const hH = 110000;        // ~0.30 cm
  const rowH = boxH * 3 + hH;
  const gapY = 120000;      // ~0.33 cm
  const stepY = rowH + gapY; // 650000

  const gapX = 220000;      // ~0.61 cm
  const stepX = nodeW + gapX; // 700000

  const placedNodes = [];
  acts.forEach(a => {
    const col = levels[a.id] || 0;
    const tr = tracks[a.id] !== undefined ? tracks[a.id] : 0;
    a.x = col * stepX;
    a.y = Math.round(tr * stepY);
    placedNodes.push(a);
  });

  // 5. Build edges
  const edges = [];
  acts.forEach(a => {
    a.pred.forEach(p => {
      const srcNode = placedNodes.find(n => n.id === p);
      if (srcNode) {
        const isCrit = a.crit && srcNode.crit;
        edges.push([p, a.id, isCrit]);
      }
    });
  });

  let maxX = 0;
  let maxY = 0;
  placedNodes.forEach(n => {
    if (n.x + nodeW > maxX) maxX = n.x + nodeW;
    if (n.y + boxH * 3 + hH > maxY) maxY = n.y + boxH * 3 + hH;
  });
  maxX += 100000;
  maxY += 100000;

  let shapeCounter = 1000;
  const shapesXml = [];
  const nodeMidShapeIds = {};

  function makeBox(sId, x, y, cx, cy, text, isCrit, isBorderless, isBold) {
    const lnXml = isBorderless
      ? '<a:ln><a:noFill/></a:ln>'
      : `<a:ln w="${isCrit ? 19000 : 12700}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`;
    const fillXml = isCrit && !isBorderless
      ? '<a:solidFill><a:srgbClr val="F0F4F8"/></a:solidFill>'
      : '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>';
    const bTag = isBold ? '<w:b/>' : '';

    return `
      <wps:wsp>
        <wps:cNvPr id="${sId}" name="B${sId}"/>
        <wps:cNvSpPr/>
        <wps:spPr>
          <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          ${fillXml}
          ${lnXml}
        </wps:spPr>
        <wps:txbx>
          <w:txbxContent>
            <w:p>
              <w:pPr>
                <w:jc w:val="center"/>
                <w:spacing w:line="180" w:lineRule="auto" w:before="0" w:after="0"/>
                <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="15"/>${bTag}</w:rPr>
              </w:pPr>
              <w:r>
                <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="15"/>${bTag}</w:rPr>
                <w:t>${text}</w:t>
              </w:r>
            </w:p>
          </w:txbxContent>
        </wps:txbx>
        <wps:bodyPr anchor="ctr" lIns="2000" rIns="2000" tIns="2000" bIns="2000"/>
      </wps:wsp>
    `;
  }

  placedNodes.forEach(n => {
    const nx = n.x;
    const ny = n.y;
    const crit = n.crit;

    shapesXml.push(makeBox(shapeCounter++, nx, ny, boxW, boxH, String(n.es), crit, false, false));
    shapesXml.push(makeBox(shapeCounter++, nx + boxW, ny, boxW, boxH, String(n.ef), crit, false, false));
    const sMid = shapeCounter++;
    nodeMidShapeIds[n.id] = sMid;
    const actLbl = n.t > 0 ? `${n.act} (t=${n.t})` : n.act;
    shapesXml.push(makeBox(sMid, nx, ny + boxH, nodeW, boxH, actLbl, crit, false, true));
    shapesXml.push(makeBox(shapeCounter++, nx, ny + boxH * 2, boxW, boxH, String(n.ls), crit, false, false));
    shapesXml.push(makeBox(shapeCounter++, nx + boxW, ny + boxH * 2, boxW, boxH, String(n.lf), crit, false, false));
    shapesXml.push(makeBox(shapeCounter++, nx + Math.floor((nodeW - 300000) / 2), ny + boxH * 3, 300000, hH, `h=${n.h}`, crit, true, crit));
  });

  edges.forEach(([src, dst, isCrit]) => {
    const srcId = nodeMidShapeIds[src];
    const dstId = nodeMidShapeIds[dst];
    const srcNode = placedNodes.find(n => n.id === src);
    const dstNode = placedNodes.find(n => n.id === dst);
    if (!srcNode || !dstNode) return;

    // Always exit from right center of source and enter left center of destination
    const stIdx = 3; // exit right center
    const endIdx = 1; // enter left center
    const x1 = srcNode.x + nodeW;
    const y1 = srcNode.y + boxH + Math.floor(boxH / 2);
    const x2 = dstNode.x;
    const y2 = dstNode.y + boxH + Math.floor(boxH / 2);

    const minX = Math.min(x1, x2);
    const minY = Math.min(y1, y2);
    const cx = Math.max(12700, Math.abs(x2 - x1));
    const cy = Math.max(12700, Math.abs(y2 - y1));
    const flipH = x2 < x1 ? ' flipH="1"' : '';
    const flipV = y2 < y1 ? ' flipV="1"' : '';

    const connId = shapeCounter++;
    const lw = isCrit ? 25000 : 12700;

    shapesXml.push(`
      <wps:wsp>
        <wps:cNvPr id="${connId}" name="F${connId}"/>
        <wps:cNvCnPr><a:stCxn id="${srcId}" idx="${stIdx}"/><a:endCxn id="${dstId}" idx="${endIdx}"/></wps:cNvCnPr>
        <wps:spPr>
          <a:xfrm${flipH}${flipV}><a:off x="${minX}" y="${minY}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="${lw}">
            <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            <a:tailEnd type="triangle" w="med" len="med"/>
          </a:ln>
        </wps:spPr>
        <wps:bodyPr/>
      </wps:wsp>
    `);
  });

  const drawingXml = `
    <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wp:inline>
        <wp:extent cx="${maxX}" cy="${maxY}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="5000" name="Red PERT Organizada"/>
        <wp:cNvGraphicFramePr><a:graphicFrameLocks/></wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp>
              <wpg:cNvGrpSpPr/>
              <wpg:grpSpPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="${maxX}" cy="${maxY}"/>
                  <a:chOff x="0" y="0"/>
                  <a:chExt cx="${maxX}" cy="${maxY}"/>
                </a:xfrm>
              </wpg:grpSpPr>
              ${shapesXml.join("")}
            </wpg:wgp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  `;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:r>
              ${drawingXml}
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

// Button: Insert Generated Red PERT Drawing
if (btnInsertDrawing) {
  btnInsertDrawing.addEventListener("click", async () => {
    if (!currentPertOoxml) {
      alert("No hay diagrama de red generado para insertar.");
      return;
    }

    if (typeof Word !== "undefined" && Word.run) {
      try {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.insertOoxml(currentPertOoxml, Word.InsertLocation.replace);
          await context.sync();
        });
        alert("¡Diagrama de red vectorial insertado exitosamente en Word! Puedes editarlo o reacomodar nodos con doble clic.");
        return;
      } catch (e) {
        console.warn("Word.run insertOoxml fallback:", e);
        alert("Error al insertar el diagrama en Word: " + e.message);
      }
    }

    if (typeof Office !== "undefined" && Office.context && Office.context.document) {
      Office.context.document.setSelectedDataAsync(currentPertOoxml, { coercionType: Office.CoercionType.Ooxml }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          alert("¡Diagrama insertado exitosamente en Word!");
        } else {
          alert("Error al insertar: " + result.error.message);
        }
      });
      return;
    }

    alert("Para insertar el dibujo, abre el complemento dentro de Word.");
  });
}

// Button: Test 3-node Drawing
const btnTestDrawing = document.getElementById("btn-test-drawing");
if (btnTestDrawing) {
  btnTestDrawing.addEventListener("click", async () => {
    try {
      const sampleNodes = [
        { id: 'INICIO', act: 'INICIO', t: 0, pred: [], es: 0, ef: 0, ls: 0, lf: 0, h: 0, crit: true },
        { id: 'A', act: 'A', t: 3, pred: ['INICIO'], es: 0, ef: 3, ls: 0, lf: 3, h: 0, crit: true },
        { id: 'C', act: 'C', t: 1, pred: ['A'], es: 3, ef: 4, ls: 3, lf: 4, h: 0, crit: true }
      ];
      const opcXml = buildFullPertOoxmlFromActivities(sampleNodes);

      if (typeof Word !== "undefined" && Word.run) {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.insertOoxml(opcXml, Word.InsertLocation.replace);
          await context.sync();
        });
        alert("¡Éxito! Lienzo de dibujo con 3 nodos insertado. Haz doble clic para editar.");
        return;
      }

      if (typeof Office !== "undefined" && Office.context && Office.context.document) {
        Office.context.document.setSelectedDataAsync(opcXml, { coercionType: Office.CoercionType.Ooxml }, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            alert("¡Éxito! Lienzo de dibujo insertado.");
          } else {
            alert("Error al insertar OOXML: " + result.error.message);
          }
        });
        return;
      }

      alert("Abre el complemento dentro de Word Online para probar la inserción.");
    } catch (err) {
      console.error("Test drawing insertion failed:", err);
      alert("Error de Word al insertar dibujo OOXML: " + (err.message || err));
    }
  });
}

// Display a specific conversation turn and update pinned headers
function showTurn(idx) {
  if (idx < 0 || idx >= chatTurns.length) return;
  activeTurnIndex = idx;
  const turn = chatTurns[idx];

  // 1. Sync active data for Word insertion and copy
  currentRawMarkdown = turn.responseMarkdown;
  currentRenderedHtml = turn.renderedHtml;
  currentPertOoxml = turn.pertOoxml;

  // 2. Update user query box
  if (turnUserText) turnUserText.textContent = turn.query;

  // 3. Render markdown HTML into resultView
  if (resultView) {
    resultView.innerHTML = turn.renderedHtml;
    resultView.scrollTop = 0;
  }

  // 4. Update PERT vector drawing button
  if (btnInsertDrawing) {
    btnInsertDrawing.style.display = turn.pertOoxml ? "block" : "none";
  }

  // 5. Update Turn Badge
  if (turnBadge) {
    turnBadge.textContent = `Consulta ${idx + 1} de ${chatTurns.length}`;
  }

  // 6. Navigation Buttons state
  if (btnTurnPrev) btnTurnPrev.disabled = (idx <= 0);
  if (btnTurnNext) btnTurnNext.disabled = (idx >= chatTurns.length - 1);

  // 7. Pinned Previous Consultation Header (Sticky above current response)
  if (pinnedHistoryBanner && pinnedPrevTitle) {
    if (idx > 0) {
      const prevTurn = chatTurns[idx - 1];
      const snippet = (prevTurn.query || "").replace(/\s+/g, " ").trim();
      const shortSnippet = snippet.length > 45 ? snippet.slice(0, 42) + "..." : snippet;
      pinnedPrevTitle.textContent = `#${idx}: "${shortSnippet}"`;
      pinnedHistoryBanner.style.display = "block";
    } else {
      pinnedHistoryBanner.style.display = "none";
    }
  }

  // 8. Pinned Next Consultation Banner (When browsing past turns)
  if (pinnedNextBanner && pinnedNextTitle) {
    if (idx < chatTurns.length - 1) {
      const nextTurn = chatTurns[idx + 1];
      const snippet = (nextTurn.query || "").replace(/\s+/g, " ").trim();
      const shortSnippet = snippet.length > 45 ? snippet.slice(0, 42) + "..." : snippet;
      pinnedNextTitle.textContent = `#${idx + 2}: "${shortSnippet}"`;
      pinnedNextBanner.style.display = "block";
    } else {
      pinnedNextBanner.style.display = "none";
    }
  }

  resultContainer.style.display = "flex";
}

// Navigation button listeners
if (btnPinnedPrev) {
  btnPinnedPrev.addEventListener("click", () => {
    if (activeTurnIndex > 0) showTurn(activeTurnIndex - 1);
  });
}
if (btnPinnedNext) {
  btnPinnedNext.addEventListener("click", () => {
    if (activeTurnIndex < chatTurns.length - 1) showTurn(activeTurnIndex + 1);
  });
}
if (btnTurnPrev) {
  btnTurnPrev.addEventListener("click", () => {
    if (activeTurnIndex > 0) showTurn(activeTurnIndex - 1);
  });
}
if (btnTurnNext) {
  btnTurnNext.addEventListener("click", () => {
    if (activeTurnIndex < chatTurns.length - 1) showTurn(activeTurnIndex + 1);
  });
}

// Solve exercise & Multi-turn Conversational Memory
btnSolve.addEventListener("click", async () => {
  const query = promptInput.value.trim();
  if (!query) {
    alert("Por favor ingresa o selecciona la consigna o tu pregunta de seguimiento.");
    return;
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    apiModal.classList.add("active");
    alert("Debes configurar tu Gemini API Key primero.");
    return;
  }

  // Format prompt based on whether it is a fresh problem or a follow-up
  let userPrompt = "";
  if (chatHistory.length === 0) {
    const type = exerciseType.value;
    userPrompt = buildOperationalResearchPrompt(query, type);
  } else {
    userPrompt = `PREGUNTA O AJUSTE DE SEGUIMIENTO SOBRE EL EJERCICIO ANTERIOR:
${query}

REGLAS DE FORMATO:
- Mantén la coherencia matemática estricta con el ejercicio anterior.
- Responde directamente al pedido con explicaciones claras o tablas actualizadas.
- Si presentas tablas, usa formato Markdown completo para Microsoft Word.
- CERO bloques de código (\`\`\`), CERO LaTeX crudo ($).
- Si se modifican actividades de la red PERT, incluye al final el bloque actualizado:
<!-- PERT_DATA: [...] -->`;
  }

  chatHistory.push({ role: "user", parts: [{ text: userPrompt }] });

  // Show loading
  btnSolve.disabled = true;
  loadingBox.classList.add("active");
  resultContainer.style.display = "none";

  try {
    const responseText = await callGeminiAPI(chatHistory, apiKey);
    chatHistory.push({ role: "model", parts: [{ text: responseText }] });

    // Switch UI to multi-turn mode
    btnSolve.textContent = "Preguntar / Ajustar";
    promptInput.value = "";
    promptInput.placeholder = "Haz una pregunta de seguimiento, pide un ajuste o toca 'Limpiar' para un nuevo ejercicio...";

    // Check if response contains PERT data to generate native Vector Drawing Canvas
    let turnPertOoxml = null;
    const pertActivities = parsePertTable(responseText);
    if (pertActivities && pertActivities.length >= 2) {
      try {
        turnPertOoxml = buildFullPertOoxmlFromActivities(pertActivities);
      } catch (e) {
        console.warn("Error building full PERT OOXML:", e);
      }
    }

    // Convert markdown to clean HTML with native Word elements
    let renderedHtml = "";
    if (typeof marked !== "undefined" && marked.parse) {
      renderedHtml = prepareWordHtml(marked.parse(responseText));
    } else {
      renderedHtml = prepareWordHtml(responseText);
    }

    // Record turn
    const newTurn = {
      index: chatTurns.length,
      query: query,
      responseMarkdown: responseText,
      renderedHtml: renderedHtml,
      pertOoxml: turnPertOoxml,
    };
    chatTurns.push(newTurn);

    // Show latest turn (automatically updates pinned previous header!)
    showTurn(chatTurns.length - 1);
  } catch (err) {
    chatHistory.pop(); // Revert user message on error so history stays valid
    alert("Error al resolver: " + err.message);
  } finally {
    btnSolve.disabled = false;
    loadingBox.classList.remove("active");
  }
});

// Specialized prompt builder for Investigación Operativa
function buildOperationalResearchPrompt(consigna, type) {
  let typeInstruction = "";

  if (type === "simplex" || consigna.toLowerCase().includes("simplex") || consigna.toLowerCase().includes("programación lineal")) {
    typeInstruction = `
TIPO DE PROBLEMA: PROGRAMACIÓN LINEAL / MÉTODO SIMPLEX.
ESTRUCTURA OBLIGATORIA DEL DESARROLLO (Paso a Paso):
1. PLANTEO FORMAL:
   - Definición explícita de variables de decisión (ej: X1 = ..., X2 = ...).
   - Función Objetivo formal (Zmax o Zmin con unidades monetarias o métricas).
   - Sistema de Restricciones canónicas con unidades.
   - Condición de no negatividad (X1, X2 >= 0).
2. FORMA ESTÁNDAR:
   - Incorporación de variables de holgura (S1, S2, etc.) o artificiales.
3. TABLAS SIMPLEX PASO A PASO (Formato de tabla Markdown completo):
   - Tabla Simplex Inicial (Tabla 0): con fila Cj, variables básicas, coeficientes, fila Zj y fila (Cj - Zj).
   - Identificación justificada de la COLUMNA PIVOTE (variable que entra) y FILA PIVOTE (cálculo de cocientes mínimos / variable que sale). Indicar claramente el ELEMENTO PIVOTE.
   - Tabla Simplex Iteración 1, Iteración 2 (si aplica) con las operaciones de Gauss-Jordan detalladas.
   - Tabla Simplex Final (condición de parada alcanzada).
4. SOLUCIÓN ÓPTIMA E INTERPRETACIÓN:
   - Valores óptimos finales de cada variable y valor de Z*.
   - Interpretación textual en el contexto del problema (ej: cantidad exacta a producir o asientos a ofrecer).`;
  } else if (type === "pert" || consigna.toLowerCase().includes("red") || consigna.toLowerCase().includes("camino crítico") || consigna.toLowerCase().includes("holgura")) {
    typeInstruction = `
TIPO DE PROBLEMA: DIAGRAMA DE REDES / PERT - CPM.
DEBES RESOLVER EL PROBLEMA SIGUIENDO ESTRICTAMENTE ESTE ORDEN DE RAZONAMIENTO MATEMÁTICO:

1. PASO HACIA ADELANTE (FORWARD PASS) - TIEMPOS TEMPRANOS:
Calcula para cada actividad su Inicio Temprano (ES) y Fin Temprano (EF = ES + t):
- Para actividades iniciales (sin predecesores): ES = 0.
- Para actividades con predecesoras: ES = max(EF de todas sus predecesoras directas).
Muestra explícitamente cada ecuación y resultado aritmético paso a paso (ej: "ES(C) = EF(A) = 6 ➔ EF(C) = 6 + 2 = 8 meses").

2. TIEMPO MÍNIMO DEL PROYECTO (T*):
Identifica todas las actividades terminales (aquellas que NO tienen sucesoras directas, es decir: F, M, N, P).
Calcula formalmente el tiempo total del proyecto: T* = max(EF de todas las terminales). Muestra la comparación: max(EF(F), EF(M), EF(N), EF(P)).

3. PASO HACIA ATRÁS (BACKWARD PASS) - TIEMPOS TARDÍOS:
(REGLA UNIVERSAL ESTRICTA: Para TODAS las actividades terminales F, M, N, P, su Fin Tardío es IGUAL a T*: LF = T*. Ninguna terminal puede tener LF distinto a T*).
Calcula para cada actividad su Fin Tardío (LF) e Inicio Tardío (LS = LF - t):
- Para actividades con sucesoras: LF = min(LS de todas sus sucesoras directas).
Muestra explícitamente cada cálculo hacia atrás.

4. CÁLCULO DE HOLGURAS TOTALES Y DETERMINACIÓN DEL CAMINO CRÍTICO:
- Calcula para cada actividad la Holgura Total: h = LS - ES = LF - EF.
- Identifica las actividades críticas con holgura nula (h = 0).
- Define la secuencia exacta del Camino Crítico principal.

5. ANÁLISIS COMPARATIVO DE RUTAS DE LA RED (TABLA DE WORD):
Genera obligatoriamente una tabla formal de Word con la comparación de todas las rutas posibles desde INICIO hasta FIN:
| Ruta | Secuencia de la Ruta | Operación de Duración | Duración Total | ¿Es Ruta Crítica? |
| :---: | :--- | :---: | :---: | :---: |
(Enumera cada ruta y resalta en negrita la ruta con mayor duración: T*).

6. TABLA MAESTRA DE ACTIVIDADES Y HOLGURAS (TABLA CUADRADA DE 9 COLUMNAS):
Vuelca los resultados matemáticos calculados anteriormente en la tabla final con bordes:
| Actividad | Predecesora | Duración (t) | Inicio Temprano (ES) | Fin Temprano (EF) | Inicio Tardío (LS) | Fin Tardío (LF) | Holgura Total (H = LF - EF) | ¿Crítica? |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |

7. ESTRUCTURA JSON OCULTA (OBLIGATORIA AL FINAL DEL TEXTO):
Al final de todo el texto, incluye un bloque HTML oculto con el resumen JSON de las actividades para alimentar el generador vectorial del dibujo:
<!-- PERT_DATA:
[
  {"id": "A", "act": "A", "t": 6, "pred": [], "es": 0, "ef": 6, "ls": 0, "lf": 6, "h": 0, "crit": true},
  ...
]
-->`;
  } else if (type === "arbol" || consigna.toLowerCase().includes("árbol") || consigna.toLowerCase().includes("probabilidades") || consigna.toLowerCase().includes("estudio de mercado")) {
    typeInstruction = `
TIPO DE PROBLEMA: TEORÍA DE LA DECISIÓN / ÁRBOLES DE DECISIÓN.
ESTRUCTURA OBLIGATORIA DEL DESARROLLO (Paso a Paso):
1. IDENTIFICACIÓN DE ALTERNATIVAS, ESTADOS DE LA NATURALEZA Y PROBABILIDADES:
   - Listado de alternativas de decisión (Nodos cuadrados).
   - Listado de eventos aleatorios con sus probabilidades asociadas (Nodos circulares).
   - Beneficios netos o costos de cada rama terminal.
2. ESQUEMA JERÁRQUICO DEL ÁRBOL:
   - Detalle rama por rama de la estructura del árbol de decisión.
3. CÁLCULO DEL VALOR ESPERADO (VE / EMV) HACIA ATRÁS:
   - Cálculo matemático paso a paso de cada nodo de probabilidad de derecha a izquierda: VE = sum(Probabilidad * Beneficio).
   - Comparación en cada nodo de decisión y poda de ramas no óptimas.
4. DECISIÓN FINAL RECOMENDADA:
   - Detalle de la mejor estrategia a seguir y el Valor Esperado óptimo resultante.`;
  } else if (type === "transporte" || consigna.toLowerCase().includes("transporte") || consigna.toLowerCase().includes("costos minimos") || consigna.toLowerCase().includes("vogel") || consigna.toLowerCase().includes("ofertas")) {
    typeInstruction = `
TIPO DE PROBLEMA: MODELO DE TRANSPORTE Y ASIGNACIÓN.
ESTRUCTURA OBLIGATORIA DEL DESARROLLO (Paso a Paso):
1. VERIFICACIÓN DE EQUILIBRIO:
   - Suma total de Oferta vs Suma total de Demanda (balanceo).
2. MATRIZ DE COSTOS UNITARIOS INICIAL:
   - Tabla Markdown con orígenes, destinos, costos unitarios, ofertas y demandas.
3. APLICACIÓN DEL MÉTODO (MÉTODO DE COSTOS MÍNIMOS O EL SOLICITADO):
   - Asignación iterativa celda por celda: identificar la celda con menor costo, asignar la máxima cantidad posible min(Oferta, Demanda), tachar fila/columna satisfecha y ajustar remanentes.
4. MATRIZ DE DISTRIBUCIÓN FINAL:
   - Tabla completa con las cantidades asignadas en cada ruta.
5. CÁLCULO DEL COSTO TOTAL (Z):
   - Sumatoria formal Z = sum(Costo_ij * Cantidad_ij) con el resultado numérico final.`;
  } else {
    typeInstruction = `
TIPO DE PROBLEMA: EXAMEN DE INGENIERÍA / SISTEMAS.
Desarrolla la respuesta con el máximo rigor académico, pasos matemáticos detallados, justificación técnica formal y tablas completas si corresponde.`;
  }

  return `OBJETIVO: Eres un profesor titular y experto en INVESTIGACIÓN OPERATIVA de Ingeniería en Sistemas.
Tu tarea es resolver la siguiente consigna de examen generando EXCLUSIVAMENTE EL DESARROLLO COMPLETO Y PASO A PASO con todas las tablas matemáticas.

${typeInstruction}

REGLAS DE FORMATO ESTRICTAS:
1. SOLO EL DESARROLLO CON LOS PASOS: No incluyas saludos, ni charlas, ni introducciones vacías como "A continuación resuelvo...". Ve directo al paso 1.
2. FORMATO NATIVO PARA MICROSOFT WORD (CERO CÓDIGO Y CERO LATEX):
   - ESTRICTAMENTE PROHIBIDO usar bloques de código (\`\`\`), líneas de consola o signos igual (===).
   - ESTRICTAMENTE PROHIBIDO usar sintaxis LaTeX cruda (\\implies, \\max, \\text, signos de dólar $). Escribe texto claro y legible: "ES = 3  →  EF = 4 meses".
   - TODAS LAS TABLAS EN FORMATO TABLA MARKDOWN: Usa tablas Markdown prolijas (con | y -) para que Word las convierta automáticamente en tablas reales editables con cuadrícula y bordes.
3. PRECISIÓN MATEMÁTICA ABSOLUTA: Verifica todas las operaciones aritméticas (sumas, multiplicaciones de matrices, holguras, valor esperado).

CONSIGNA DEL EXAMEN:
${consigna}`;
}

// Call Google Gemini API (Single prompt or Multi-turn History)
async function callGeminiAPI(promptOrContents, apiKey) {
  const models = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-flash-latest"];
  let lastError = null;

  const contentsPayload = Array.isArray(promptOrContents)
    ? promptOrContents
    : [{ parts: [{ text: promptOrContents }] }];

  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
      const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-goog-api-key": apiKey,
        },
        body: JSON.stringify({
          contents: contentsPayload,
          generationConfig: {
            temperature: 0.0,
            maxOutputTokens: 8192,
          },
        }),
      });

      const data = await response.json();
      if (data.error) {
        lastError = data.error.message;
        continue;
      }

      const text = data.candidates?.[0]?.content?.parts?.map(p => p.text).filter(Boolean).join("\n\n");
      if (text) return text;
    } catch (err) {
      lastError = err.message;
    }
  }

  throw new Error(lastError || "No se pudo obtener respuesta de la IA.");
}
