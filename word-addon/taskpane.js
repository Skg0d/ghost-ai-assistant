// Ghost AI - Word Add-in Controller (Stealth Edition)
const STORAGE_KEY = "ghost_gemini_api_key";
const STORAGE_THEME = "ghost_addon_theme";
const STORAGE_OPACITY = "ghost_addon_opacity";

let currentRawMarkdown = "";
let currentRenderedHtml = "";

// Elements
const stealthContent = document.getElementById("stealth-content");
const panicZone = document.getElementById("panic-zone");
const promptInput = document.getElementById("prompt-input");
const exerciseType = document.getElementById("exercise-type");
const btnSolve = document.getElementById("btn-solve");
const btnClear = document.getElementById("btn-clear");
const btnGetSelection = document.getElementById("btn-get-selection");
const btnInsertWord = document.getElementById("btn-insert-word");
const btnCopyClipboard = document.getElementById("btn-copy-clipboard");
const resultContainer = document.getElementById("result-container");
const resultView = document.getElementById("result-view");
const loadingBox = document.getElementById("loading");

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

// Clear input
btnClear.addEventListener("click", () => {
  promptInput.value = "";
  resultContainer.style.display = "none";
  currentRawMarkdown = "";
  currentRenderedHtml = "";
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
// Test Drawing Generator (Vector Word Canvas)
// ==========================================
function generatePertOoxml(nodes, edges) {
  const nodeW = 733963;
  const boxW = 366982;
  const boxH = 301307;
  const hH = 250000;

  let maxX = 0;
  let maxY = 0;
  nodes.forEach(n => {
    if (n.x + nodeW > maxX) maxX = n.x + nodeW;
    if (n.y + boxH * 3 + hH > maxY) maxY = n.y + boxH * 3 + hH;
  });
  maxX += 200000;
  maxY += 200000;

  let shapeCounter = 1000;
  const shapesXml = [];
  const nodeMidShapeIds = {};

  function makeTextbox(shapeId, x, y, cx, cy, text, isBorderless, isBold) {
    const lnXml = isBorderless
      ? '<a:ln><a:noFill/></a:ln>'
      : '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';
    const bAttr = isBold ? '<w:b/>' : '';

    return `
      <wps:wsp>
        <wps:cNvPr id="${shapeId}" name="Box ${shapeId}"/>
        <wps:cNvSpPr/>
        <wps:spPr>
          <a:xfrm>
            <a:off x="${x}" y="${y}"/>
            <a:ext cx="${cx}" cy="${cy}"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
          ${lnXml}
        </wps:spPr>
        <wps:txbx>
          <w:txbxContent>
            <w:p>
              <w:pPr>
                <w:jc w:val="center"/>
                <w:spacing w:line="220" w:lineRule="auto" w:before="0" w:after="0"/>
                <w:rPr>
                  <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                  <w:sz w:val="18"/>
                  ${bAttr}
                </w:rPr>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                  <w:sz w:val="18"/>
                  ${bAttr}
                </w:rPr>
                <w:t>${text}</w:t>
              </w:r>
            </w:p>
          </w:txbxContent>
        </wps:txbx>
        <wps:bodyPr anchor="ctr" lIns="10000" rIns="10000" tIns="10000" bIns="10000"/>
      </wps:wsp>
    `;
  }

  nodes.forEach(n => {
    const nx = n.x;
    const ny = n.y;

    const sEs = shapeCounter++;
    shapesXml.push(makeTextbox(sEs, nx, ny, boxW, boxH, String(n.es)));

    const sEf = shapeCounter++;
    shapesXml.push(makeTextbox(sEf, nx + boxW, ny, boxW, boxH, String(n.ef)));

    const sMid = shapeCounter++;
    nodeMidShapeIds[n.id] = sMid;
    const actLabel = (n.t !== undefined && n.t !== null) ? `${n.act} (t=${n.t})` : n.act;
    shapesXml.push(makeTextbox(sMid, nx, ny + boxH, nodeW, boxH, actLabel, false, true));

    const sLs = shapeCounter++;
    shapesXml.push(makeTextbox(sLs, nx, ny + boxH * 2, boxW, boxH, String(n.ls)));

    const sLf = shapeCounter++;
    shapesXml.push(makeTextbox(sLf, nx + boxW, ny + boxH * 2, boxW, boxH, String(n.lf)));

    const sH = shapeCounter++;
    const hX = nx + Math.floor((nodeW - 454122) / 2);
    shapesXml.push(makeTextbox(sH, hX, ny + boxH * 3, 454122, hH, `h=${n.h}`, true, false));
  });

  edges.forEach(([src, dst]) => {
    const srcId = nodeMidShapeIds[src];
    const dstId = nodeMidShapeIds[dst];
    const srcNode = nodes.find(n => n.id === src);
    const dstNode = nodes.find(n => n.id === dst);
    if (!srcNode || !dstNode) return;

    const x1 = srcNode.x + nodeW;
    const y1 = srcNode.y + boxH + Math.floor(boxH / 2);
    const x2 = dstNode.x;
    const y2 = dstNode.y + boxH + Math.floor(boxH / 2);

    const cx = x2 - x1;
    const cy = y2 - y1;
    const connId = shapeCounter++;

    shapesXml.push(`
      <wps:wsp>
        <wps:cNvPr id="${connId}" name="Flecha ${connId}"/>
        <wps:cNvCnPr>
          <a:stCxn id="${srcId}" idx="3"/>
          <a:endCxn id="${dstId}" idx="1"/>
        </wps:cNvCnPr>
        <wps:spPr>
          <a:xfrm>
            <a:off x="${x1}" y="${y1}"/>
            <a:ext cx="${cx}" cy="${cy}"/>
          </a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="15000">
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
        <wp:docPr id="1000" name="Lienzo PERT"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks/>
        </wp:cNvGraphicFramePr>
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

const btnTestDrawing = document.getElementById("btn-test-drawing");
if (btnTestDrawing) {
  btnTestDrawing.addEventListener("click", async () => {
    try {
      const sampleNodes = [
        { id: 'INICIO', act: 'INICIO', t: 0, es: 0, ef: 0, ls: 0, lf: 0, h: 0, x: 0, y: 150000 },
        { id: 'A', act: 'A', t: 3, es: 0, ef: 3, ls: 0, lf: 3, h: 0, x: 1400000, y: 150000 },
        { id: 'C', act: 'C', t: 1, es: 3, ef: 4, ls: 3, lf: 4, h: 0, x: 2800000, y: 150000 }
      ];
      const sampleEdges = [
        ['INICIO', 'A'],
        ['A', 'C']
      ];
      const opcXml = generatePertOoxml(sampleNodes, sampleEdges);

      if (typeof Word !== "undefined" && Word.run) {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.insertOoxml(opcXml, Word.InsertLocation.replace);
          await context.sync();
        });
        alert("¡Éxito! Lienzo de dibujo con 3 nodos y flechas insertado en tu Word. Haz clic sobre él en Word para comprobar si se activa el editor de dibujo.");
        return;
      }

      if (typeof Office !== "undefined" && Office.context && Office.context.document) {
        Office.context.document.setSelectedDataAsync(opcXml, { coercionType: Office.CoercionType.Ooxml }, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            alert("¡Éxito! Lienzo de dibujo insertado en Word.");
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

// Solve exercise
btnSolve.addEventListener("click", async () => {
  const query = promptInput.value.trim();
  if (!query) {
    alert("Por favor ingresa o selecciona la consigna del ejercicio a resolver.");
    return;
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    apiModal.classList.add("active");
    alert("Debes configurar tu Gemini API Key primero.");
    return;
  }

  // Show loading
  btnSolve.disabled = true;
  loadingBox.classList.add("active");
  resultContainer.style.display = "none";

  const type = exerciseType.value;
  const prompt = buildOperationalResearchPrompt(query, type);

  try {
    const responseText = await callGeminiAPI(prompt, apiKey);
    currentRawMarkdown = responseText;

    // Convert markdown to clean HTML with native Word elements
    if (typeof marked !== "undefined" && marked.parse) {
      currentRenderedHtml = prepareWordHtml(marked.parse(responseText));
    } else {
      currentRenderedHtml = prepareWordHtml(responseText);
    }

    resultView.innerHTML = currentRenderedHtml;
    resultContainer.style.display = "flex";
  } catch (err) {
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
DEBES GENERAR EXACTAMENTE LAS SIGUIENTES SECCIONES Y TABLAS PARA MICROSOFT WORD:

1. DIAGRAMA DE RED (CUADROS Y FLECHAS DE WORD):
ESTRICTAMENTE PROHIBIDO USAR BLOQUES DE CÓDIGO (\`\`\`), LÍNEAS DE CONSOLA O SIGNOS IGUAL (===).
Para que Word dibuje cuadros reales con bordes y flechas nativas:
A. Representa cada rama de la red como una secuencia de nodos encuadrados entre corchetes [ ] conectados por flechas ➔:
   - Ruta Principal (Crítica): [ INICIO ] ➔ [ Actividad (duración) ] ➔ ... ➔ [ FIN: T* ]
   - Demás Rutas: [ INICIO ] ➔ ... ➔ [ FIN: duración ]
   (Cada corchete [ ... ] se transformará automáticamente en un cuadro con borde negro sólido en Word).

B. Tabla de Conexiones de Nodos de la Red (AON):
   Genera una tabla cuadrada de Word que sintetice la estructura de la red:
   | Nodo / Actividad | Duración (t) | Predecesoras Directas | Conecta hacia (Sucesoras) |
   | :---: | :---: | :---: | :---: |

2. ANÁLISIS COMPARATIVO DE RUTAS DE LA RED (TABLA DE WORD):
Genera obligatoriamente una tabla formal de Word con la comparación de todas las rutas posibles desde (INICIO) hasta (FIN):
| Ruta | Secuencia de la Ruta | Operación de Duración | Duración Total | ¿Es Ruta Crítica? |
| :---: | :--- | :---: | :---: | :---: |
(Enumera cada ruta y resalta en negrita la ruta con mayor duración: T*).

3. CÁLCULO DE TIEMPOS TEMPRANOS Y TARDÍOS (SIN CÓDIGO NI LATEX):
CERO signos de dólar $, CERO comandos LaTeX como \\implies o \\text{}.
Escribe las ecuaciones en texto limpio y natural para Word:
- Paso hacia adelante (Forward Pass):
  * ES = 0 para actividades iniciales.
  * EF = ES + t.
  * ES = max(EF de predecesoras) para convergencias.
  * Tiempo mínimo del proyecto: T* = max(EF de terminales).
  (Detalla los cálculos de cada actividad: ej: ES(C) = EF(A) = 3  →  EF(C) = 3 + 1 = 4 meses).
- Paso hacia atrás (Backward Pass) - REGLA UNIVERSAL ESTRICTA:
  * OBLIGATORIO: Para TODAS las actividades terminales (sin sucesoras), su Tiempo de Finalización Tardía es IGUAL a T* (LF = T*). Ninguna actividad puede tener LF > T*.
  * LS = LF - t.
  * LF = min(LS de sucesoras directas) para bifurcaciones.
  (Detalla los cálculos paso a paso).

4. TABLA MAESTRA DE ACTIVIDADES Y HOLGURAS (TABLA CUADRADA DE 9 COLUMNAS):
Genera obligatoriamente la tabla con esta estructura exacta de 9 columnas, verificando cada cálculo aritmético celda por celda:
| Actividad | Predecesora | Duración (t) | Inicio Temprano (ES) | Fin Temprano (EF) | Inicio Tardío (LS) | Fin Tardío (LF) | Holgura Total (H = LF - EF) | ¿Crítica? |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |

5. CAMINO CRÍTICO Y CONCLUSIONES:
- Duración mínima de realización del proyecto: T* (con unidad de tiempo).
- Ruta Crítica: secuencia exacta de actividades con Holgura = 0.`;
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

// Call Google Gemini API
async function callGeminiAPI(prompt, apiKey) {
  const models = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-flash-latest"];
  let lastError = null;

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
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.15,
            maxOutputTokens: 4096,
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
