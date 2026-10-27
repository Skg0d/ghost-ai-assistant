// Ghost AI - Word Add-in Controller
const STORAGE_KEY = "ghost_gemini_api_key";
const STORAGE_THEME = "ghost_addon_theme";
const STORAGE_OPACITY = "ghost_addon_opacity";

let currentRawMarkdown = "";
let currentRenderedHtml = "";

// Elements
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

// Theme and Opacity elements
const btnTheme = document.getElementById("btn-theme");
const opacitySlider = document.getElementById("opacity-slider");
const opacityVal = document.getElementById("opacity-val");

// Settings Modal elements
const btnSettings = document.getElementById("btn-settings");
const apiModal = document.getElementById("api-modal");
const modalKeyInput = document.getElementById("modal-key-input");
const btnModalSave = document.getElementById("btn-modal-save");
const btnModalCancel = document.getElementById("btn-modal-cancel");

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Ghost AI loaded inside Word host.");
  }
});

// Theme Management (Monochrome Blanco y Negro: Claro / Oscuro)
function applyTheme(theme) {
  if (theme === "dark") {
    document.body.classList.remove("theme-light");
    document.body.classList.add("theme-dark");
    btnTheme.textContent = "Modo B/N: Oscuro";
  } else {
    document.body.classList.remove("theme-dark");
    document.body.classList.add("theme-light");
    btnTheme.textContent = "Modo B/N: Claro";
  }
  localStorage.setItem(STORAGE_THEME, theme);
}

btnTheme.addEventListener("click", () => {
  const isDark = document.body.classList.contains("theme-dark");
  applyTheme(isDark ? "light" : "dark");
});

// Opacity Management
function applyOpacity(val) {
  const clamped = Math.max(20, Math.min(100, parseInt(val, 10) || 100));
  document.body.style.opacity = (clamped / 100).toString();
  opacityVal.textContent = clamped + "%";
  opacitySlider.value = clamped;
  localStorage.setItem(STORAGE_OPACITY, clamped);
}

opacitySlider.addEventListener("input", (e) => {
  applyOpacity(e.target.value);
});

// Load saved preferences
const savedTheme = localStorage.getItem(STORAGE_THEME) || "light";
applyTheme(savedTheme);

const savedOpacity = localStorage.getItem(STORAGE_OPACITY) || "100";
applyOpacity(savedOpacity);

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

// Insert solution directly into Word document
btnInsertWord.addEventListener("click", async () => {
  if (!currentRenderedHtml) {
    alert("No hay desarrollo generado para insertar.");
    return;
  }

  // Academic monochrome styling for Word document
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

    // Convert markdown to clean HTML
    if (typeof marked !== "undefined" && marked.parse) {
      currentRenderedHtml = marked.parse(responseText);
    } else {
      currentRenderedHtml = `<pre style="white-space: pre-wrap;">${responseText}</pre>`;
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
ESTRUCTURA OBLIGATORIA DEL DESARROLLO (Paso a Paso):
1. ANÁLISIS DE PREDECESORAS Y CAMINOS:
   - Identificación de todas las rutas posibles desde el inicio hasta el fin y cálculo de la duración de cada camino.
2. CÁLCULO DE TIEMPOS TEMPRANOS Y TARDÍOS:
   - Paso hacia adelante (Forward Pass): Tiempos de inicio temprano (ES) y finalización temprana (EF).
   - Paso hacia atrás (Backward Pass): Tiempos de finalización tardía (LF) e inicio tardío (LS).
3. TABLA GENERAL DE ACTIVIDADES Y HOLGURAS (Tabla Markdown obligatoria con las siguientes columnas):
   | Actividad | Predecesora | Duración (t) | Inicio Temprano (ES) | Fin Temprano (EF) | Inicio Tardío (LS) | Fin Tardío (LF) | Holgura Total (H = LF - EF) | ¿Crítica? |
4. CAMINO CRÍTICO Y DURACIÓN DEL PROYECTO:
   - Identificación explícita de las actividades con Holgura = 0.
   - Ruta del Camino Crítico (ej: A -> C -> H -> I -> N...).
   - Tiempo mínimo de realización del proyecto.`;
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
2. TODAS LAS TABLAS EN FORMATO TABLA MARKDOWN: Usa tablas Markdown prolijas (con | y -) para que Word las convierta automáticamente en tablas reales editables.
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
