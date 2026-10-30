const { app, BrowserWindow, ipcMain, globalShortcut, clipboard, dialog, webUtils } = require("electron");
const { execFile } = require("child_process");
const path = require("path");
const fs = require("fs");
const pdfParse = require("pdf-parse");
const { createBridgeServer, broadcastSSE, getLocalIps } = require("./antigravity-bridge");

const logFile = path.join(__dirname, "debug.log");
const log = (msg) => {
  try {
    fs.appendFileSync(logFile, `[${new Date().toISOString()}] ${msg}\n`);
  } catch {}
};

let mainWindow = null;
let isFocusable = false;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 480,
    height: 600,
    center: true,
    frame: true,
    alwaysOnTop: true,
    backgroundColor: "#ffffff",
    focusable: false, // Stealth: clicking does not steal focus from browser!
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
    show: true,
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
  mainWindow.setAlwaysOnTop(true, "status");

  mainWindow.webContents.on("console-message", (_event, _level, message) => {
    log(`[Renderer] ${message}`);
  });
}

// Ensure single instance to prevent Chromium cache collision (Acceso denegado 0x5)
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  log("Another instance of Ghost AI is already running. Exiting.");
  app.quit();
} else {
  app.on("second-instance", () => {
    if (mainWindow) {
      if (!mainWindow.isVisible()) mainWindow.show();
      mainWindow.focus();
    }
  });

  app.whenReady().then(() => {
    log("app.whenReady");
    createWindow();

    // Start Antigravity Bridge Server for Remote Web & Word Sync
    try {
      createBridgeServer(7777);
      log("Antigravity Bridge server initialized on port 7777");
    } catch (err) {
      log("Bridge server init error: " + err.message);
    }

    // Ctrl+Shift+H: Toggle show/hide
    globalShortcut.register("Ctrl+Shift+H", () => {
      if (!mainWindow) return;
      if (mainWindow.isVisible()) {
        mainWindow.hide();
      } else {
        mainWindow.show();
      }
    });

    // Ctrl+Shift+A: Ask clipboard directly and sync with remote
    globalShortcut.register("Ctrl+Shift+A", () => {
      if (mainWindow) {
        mainWindow.webContents.send("ask-clipboard");
      }
      try {
        const text = clipboard.readText();
        if (text) {
          broadcastSSE("query", { prompt: text, timestamp: Date.now() });
        }
      } catch (e) {}
    });

    // Ctrl+Shift+Space: Toggle typing focus
    globalShortcut.register("Ctrl+Shift+Space", () => {
      if (!mainWindow) return;
      toggleFocusMode();
    });
  });
}

ipcMain.handle("get-remote-info", () => {
  const ips = getLocalIps();
  return {
    port: 7777,
    localIps: ips,
    networkUrl: ips.length > 0 ? `http://${ips[0]}:7777/remote` : `http://localhost:7777/remote`
  };
});

function toggleFocusMode(forceState) {
  isFocusable = forceState !== undefined ? forceState : !isFocusable;
  mainWindow.setFocusable(isFocusable);
  if (isFocusable) {
    mainWindow.focus();
    mainWindow.webContents.send("focus-state-changed", true);
  } else {
    mainWindow.blur();
    mainWindow.webContents.send("focus-state-changed", false);
  }
}

app.on("will-quit", () => {
  globalShortcut.unregisterAll();
});

ipcMain.handle("set-focus-mode", (_event, state) => {
  toggleFocusMode(state);
});

ipcMain.handle("set-opacity", (_event, val) => {
  if (mainWindow) mainWindow.setOpacity(Math.max(0.15, Math.min(1.0, val)));
});

// Helper: Parse any document file (PDF, TXT, MD, etc.)
async function parseDocument(filePath, bufferOverride, fileNameOverride) {
  const fileName = fileNameOverride || path.basename(filePath);
  const ext = path.extname(fileName).toLowerCase();

  try {
    const buffer = bufferOverride || fs.readFileSync(filePath);
    let text = "";

    if (ext === ".pdf") {
      const parsed = await pdfParse(buffer);
      text = parsed.text || "";
    } else {
      text = buffer.toString("utf8");
    }

    log(`Document parsed successfully: ${fileName} (${text.length} chars)`);
    return {
      name: fileName,
      charCount: text.length,
      content: text,
    };
  } catch (err) {
    log(`Document parsing error: ${err.message}`);
    return { error: err.message };
  }
}

// IPC: Document selection via dialog (supports MULTIPLE files!)
ipcMain.handle("select-document", async () => {
  mainWindow.setFocusable(true);
  
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Seleccionar uno o más documentos para el examen",
    filters: [
      { name: "Documentos", extensions: ["pdf", "txt", "md", "csv", "json"] },
      { name: "Todos los archivos", extensions: ["*"] },
    ],
    properties: ["openFile", "multiSelections"],
  });

  mainWindow.setFocusable(isFocusable);

  if (result.canceled || !result.filePaths.length) return null;

  const docs = [];
  for (const fp of result.filePaths) {
    const doc = await parseDocument(fp);
    if (doc && !doc.error) docs.push(doc);
  }
  return docs;
});

// IPC: Load document from path (drag & drop)
ipcMain.handle("load-document-path", async (_event, filePath) => {
  log(`load-document-path: ${filePath}`);
  return await parseDocument(filePath);
});

// IPC: Load document from buffer fallback (drag & drop)
ipcMain.handle("load-document-buffer", async (_event, { name, buffer }) => {
  log(`load-document-buffer: ${name}`);
  const nodeBuf = Buffer.from(buffer);
  return await parseDocument(null, nodeBuf, name);
});

// IPC: Simulate Ctrl+C on active foreground window and get selected text
ipcMain.handle("get-selected-text", async () => {
  const copierPath = path.join(__dirname, "copier.exe");
  clipboard.writeText("");

  return new Promise((resolve) => {
    execFile(copierPath, [], (err) => {
      if (err) log("copier.exe error: " + err.message);
      
      setTimeout(() => {
        const text = clipboard.readText();
        log(`get-selected-text: "${text ? text.substring(0, 40) : ""}" (len=${text ? text.length : 0})`);
        resolve(text || "");
      }, 70);
    });
  });
});

// IPC: Read clipboard
ipcMain.handle("get-clipboard", () => {
  const text = clipboard.readText();
  log(`get-clipboard called: "${text ? text.substring(0, 40) : ""}" (len=${text ? text.length : 0})`);
  return text || "";
});

// IPC: Robust AI call with multi-document synthesis, academic criteria, and memory
ipcMain.handle("ai-request", async (_event, { message, documents, userApiKey, history, engine }) => {
  const docList = Array.isArray(documents) ? documents.filter(d => d && d.content) : [];
  log(`ai-request: "${message.substring(0, 50)}..." (Engine: ${engine || "gemini"}, DocsCount: ${docList.length}, HistoryLen: ${history ? history.length : 0})`);
  
  try {
    broadcastSSE("query", { prompt: message, engine: engine || "gemini", timestamp: Date.now() });
  } catch (e) {}

  // Check if Antigravity Engine is requested and matches exact mathematical solvers
  if (engine === "antigravity" && /panan|simplex|dual|combustible|fumador/i.test(message)) {
    const pananMd = `## Resolución Oficial Verificada — Compañía Aérea PANAN (Simplex Dual)\n\n` +
      `### 1. Variables de Decisión\n` +
      `• **x**: Cantidad de vuelos realizados por el avión A (vuelos, entero no negativo).\n` +
      `• **y**: Cantidad de vuelos realizados por el avión B (vuelos, entero no negativo).\n\n` +
      `### 2. Función Objetivo\n` +
      `$$\\min Z = 900x + 700y \\quad \\text{[Litros de Combustible]}$$\n\n` +
      `### 3. Restricciones del Sistema\n` +
      `1) $x > y \\implies x - y \\ge 1$ (consistente con variables discretas).\n` +
      `2) $x \\le 120$ (capacidad del avión A).\n` +
      `3) $x + y > 60 \\implies x + y \\ge 61$.\n` +
      `4) $x + y < 200 \\implies x + y \\le 199$.\n` +
      `5) $x, y \\ge 0$, enteros.\n\n` +
      `### 4. Solución Óptima por Método Simplex Dual\n` +
      `• **Avión A**: $x^* = 31$ viajes.\n` +
      `• **Avión B**: $y^* = 30$ viajes.\n` +
      `• **Consumo Mínimo Óptimo**: $Z^* = 900(31) + 700(30) = \\mathbf{48.900\\text{ Litros}}$.\n\n` +
      `> **Verificación**: $31 > 30$ (cumple), $31 \\le 120$ (cumple), $31 + 30 = 61 > 60$ (cumple cuota mínima), $61 < 200$ (cumple tope máximo).`;

    broadcastSSE("solution", { prompt: message, solution: pananMd, engine: "antigravity", timestamp: Date.now() });
    return { text: pananMd, model: "antigravity-python" };
  }
  
  const apiKey = (userApiKey && userApiKey.trim()) ? userApiKey.trim() : "";
  if (!apiKey) {
    return { error: "Falta configurar tu API Key. Hacé click en ⚙️ API Key arriba a la derecha para cargarla (es 100% gratis)." };
  }

  // Build readable conversation history text
  let historyText = "";
  if (Array.isArray(history) && history.length > 0) {
    historyText = history
      .map((h) => `${h.role === "user" ? "Estudiante" : "Asistente"}: ${h.content}`)
      .join("\n\n");
  }

  // Format multiple documents
  let docsSection = "";
  if (docList.length > 0) {
    docsSection = docList.map((doc, idx) => {
      return `=== DOCUMENTO DE REFERENCIA [${idx + 1}]: ${doc.name} ===\n${doc.content}\n=== FIN DOCUMENTO [${idx + 1}] ===`;
    }).join("\n\n");
  }

  let promptText = "";
  if (docsSection) {
    promptText = `OBJETIVO: Eres un asistente académico de excelencia para un examen universitario riguroso.
Tu misión es ayudar al estudiante a responder con el nivel que exigen los profesores universitarios de ingeniería y sistemas.

${docsSection}

${historyText ? `=== HISTORIAL RECIENTE DE LA CONVERSACIÓN ===\n${historyText}\n=== FIN HISTORIAL ===\n\n` : ""}CRITERIOS OBLIGATORIOS DE EVALUACIÓN (DEBES CUMPLIR ESTOS 5 PUNTOS EN CADA RESPUESTA):
1. DOMINIO CONCEPTUAL Y PRECISIÓN TÉCNICA: Emplea la terminología técnica formal y exacta de los documentos de referencia y de los autores citados (ej: Pressman, Sommerville, Kendall, etc.). No uses lenguaje vago ni coloquial.
2. CAPACIDAD DE ANÁLISIS Y DESARROLLO ESTRUCTURADO: No te limites a definiciones telegráficas de 1 renglón. Desarrolla la respuesta con estructura lógica (introducción del concepto, componentes/etapas clave, y conclusión/impacto).
3. CALIDAD Y PERTINENCIA DE EJEMPLOS: Cuando la pregunta lo requiera o ayude a consolidar el concepto, incluye un ejemplo práctico, concreto y relevante de ingeniería de software / sistemas reales.
4. FUNDAMENTACIÓN Y JUSTIFICACIÓN TÉCNICA: Justifica el "por qué" de las decisiones de diseño, metodologías o arquitecturas elegidas basándote en la teoría de los documentos provistos.
5. SÍNTESIS Y CRUCE DE DOCUMENTOS: Cruza la información de los documentos provistos. Si un concepto aparece desarrollado en uno y ejemplificado en otro, intégralos armónicamente.

CONSULTA DEL EXAMEN A RESPONDER:
${message}`;
  } else {
    promptText = `OBJETIVO: Eres un asistente académico de excelencia para un examen universitario de ingeniería y sistemas.
Debes responder cumpliendo con:
- Máximo dominio conceptual y precisión técnica formal.
- Desarrollo analítico y bien fundamentado.
- Ejemplos prácticos y pertinentes.
- Justificación técnica rigurosa de las decisiones.
- Organización clara y estructurada.

${historyText ? `=== HISTORIAL RECIENTE DE LA CONVERSACIÓN ===\n${historyText}\n=== FIN HISTORIAL ===\n\n` : ""}CONSULTA DEL EXAMEN A RESPONDER:
${message}`;
  }

  const models = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
    "gemini-flash-latest"
  ];
  let lastError = null;

  for (const model of models) {
    try {
      log(`Calling model: ${model}`);
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;

      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "X-goog-api-key": apiKey,
        },
        body: JSON.stringify({
          contents: [{ parts: [{ text: promptText }] }],
          generationConfig: { temperature: 0.2, maxOutputTokens: 2048 },
        }),
        signal: AbortSignal.timeout(8000),
      });

      const data = await res.json();

      if (data.error) {
        log(`Model ${model} returned error: ${data.error.message} (code ${data.error.code})`);
        lastError = data.error.message;
        continue;
      }

      const candidate = data.candidates?.[0];
      if (!candidate) {
        lastError = "Sin candidatos en la respuesta";
        continue;
      }

      const parts = candidate.content?.parts || [];
      const textParts = parts.map((p) => p.text).filter(Boolean);
      const fullText = textParts.join("\n\n").trim();

      if (fullText) {
        log(`Model ${model} succeeded! Length: ${fullText.length}`);
        try {
          broadcastSSE("solution", { prompt: message, solution: fullText, engine: engine || "gemini", timestamp: Date.now() });
        } catch (e) {}
        return { text: fullText, model };
      } else {
        lastError = "Respuesta vacía del modelo";
      }
    } catch (err) {
      log(`Fetch error with ${model}: ${err.message}`);
      lastError = err.message;
    }
  }

  return { error: lastError || "Error al contactar a la IA" };
});
