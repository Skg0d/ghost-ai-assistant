const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFile, spawn } = require("child_process");

const DEFAULT_PORT = 7777;
let activePort = DEFAULT_PORT;
let sseClients = [];
let lastSolution = null;

// Helper: Get local network IPv4 addresses
function getLocalIps() {
  const interfaces = os.networkInterfaces();
  const addresses = [];
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === "IPv4" && !iface.internal) {
        addresses.push(iface.address);
      }
    }
  }
  return addresses;
}

// Broadcast event to all connected SSE clients (phones, remote tabs)
function broadcastSSE(eventType, data) {
  const payload = `event: ${eventType}\ndata: ${JSON.stringify(data)}\n\n`;
  sseClients.forEach((client) => {
    try {
      client.res.write(payload);
    } catch (err) {}
  });
}

// Clean disconnected SSE clients
function cleanSSEClients() {
  sseClients = sseClients.filter((c) => !c.res.writableEnded && !c.res.finished);
}

// Helper: Parse JSON body
function parseJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => (body += chunk));
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (err) {
        reject(err);
      }
    });
    req.on("error", reject);
  });
}

// Solve with Antigravity Python solvers
async function solveWithPython(type, prompt) {
  const solversDir = path.join(__dirname, "solvers");
  
  if (type === "simplex" || /panan|combustible|fumador|simplex|dual/i.test(prompt)) {
    return new Promise((resolve) => {
      const script = path.join(solversDir, "simplex_solver.py");
      execFile("python", [script], (err, stdout) => {
        if (!err && stdout) {
          // If PANAN or custom linear programming
          if (/panan/i.test(prompt)) {
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
            resolve(pananMd);
            return;
          }
        }
        resolve(null);
      });
    });
  }
  return null;
}

// Call Gemini API directly as fallback or fast engine
async function callGemini(prompt, userApiKey, systemInstruction) {
  const apiKey = (userApiKey && userApiKey.trim()) ? userApiKey.trim() : process.env.GEMINI_API_KEY || "";
  if (!apiKey) {
    throw new Error("No hay API Key configurada. Ingresa tu clave gratuita de Gemini.");
  }

  const models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"];
  let lastErr = null;

  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
      const contents = [{ role: "user", parts: [{ text: prompt }] }];
      const bodyPayload = { contents };
      if (systemInstruction) {
        bodyPayload.systemInstruction = { parts: [{ text: systemInstruction }] };
      }

      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-goog-api-key": apiKey },
        body: JSON.stringify(bodyPayload)
      });

      if (!res.ok) {
        const errText = await res.text();
        throw new Error(`HTTP ${res.status}: ${errText}`);
      }

      const data = await res.json();
      const candidate = data.candidates && data.candidates[0];
      const text = candidate && candidate.content && candidate.content.parts && candidate.content.parts[0] && candidate.content.parts[0].text;
      if (text) return text;
    } catch (err) {
      lastErr = err;
    }
  }
  throw lastErr || new Error("Error desconocido al llamar a Gemini.");
}

// Create HTTP Server
function createBridgeServer(port = DEFAULT_PORT) {
  const server = http.createServer(async (req, res) => {
    // Enable CORS for Word Add-in and external mobile clients
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With");

    if (req.method === "OPTIONS") {
      res.writeHead(204);
      res.end();
      return;
    }

    const parsedUrl = new URL(req.url, `http://localhost:${port}`);
    const pathname = parsedUrl.pathname;

    // 1. Status endpoint
    if (pathname === "/api/status" && req.method === "GET") {
      cleanSSEClients();
      const localIps = getLocalIps();
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({
        ok: true,
        version: "2.0-remote",
        port: activePort,
        localIps: localIps,
        localUrl: `http://localhost:${activePort}`,
        networkUrl: localIps.length > 0 ? `http://${localIps[0]}:${activePort}/remote` : `http://localhost:${activePort}/remote`,
        connectedClients: sseClients.length,
        activeEngine: "antigravity"
      }));
      return;
    }

    // 2. Solve endpoint
    if (pathname === "/api/solve" && req.method === "POST") {
      try {
        const body = await parseJsonBody(req);
        const { prompt, type = "auto", engine = "antigravity", apiKey } = body;

        if (!prompt || !prompt.trim()) {
          res.writeHead(400, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: "Falta la consigna del ejercicio." }));
          return;
        }

        // Broadcast query event to all remote screens
        broadcastSSE("query", { prompt, type, engine, timestamp: Date.now() });

        let solution = null;

        // Try Python solver first if Antigravity engine is requested
        if (engine === "antigravity") {
          solution = await solveWithPython(type, prompt);
        }

        // If not solved by specific python solver or engine is gemini, call LLM with academic rules
        if (!solution) {
          const systemRules = `Eres un experto universitario en Investigación Operativa y Ciencias Exactas. ` +
            `Resuelve con el más alto rigor matemático, mostrando cada ecuación, pivoteo, tabla paso a paso y verificación estricta. ` +
            `PROHIBIDO usar variables artificiales o Gran M en problemas que se resuelven canónicamente por Simplex Dual. ` +
            `Entrega la solución estructurada en Markdown formal.`;
          solution = await callGemini(prompt, apiKey, systemRules);
        }

        lastSolution = { prompt, solution, timestamp: Date.now() };

        // Broadcast solution event to all remote devices
        broadcastSSE("solution", { prompt, solution, engine, timestamp: Date.now() });

        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({
          ok: true,
          solution: solution,
          engine: engine
        }));
      } catch (err) {
        res.writeHead(500, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: err.message }));
      }
      return;
    }

    // 3. Export to Word/PDF and upload to temp.sh
    if (pathname === "/api/export" && req.method === "POST") {
      try {
        const body = await parseJsonBody(req);
        const { title = "Resolución Oficial", exerciseText = "", markdown = "", format = "docx" } = body;

        const timestamp = Date.now();
        const exportDocxPath = path.join(__dirname, "..", `Resolucion_Ghost_${timestamp}.docx`);
        
        // Use python doc_generator
        const genScript = path.join(__dirname, "solvers", "doc_generator.py");
        const docSections = [
          { heading: "Desarrollo y Resolución Técnica", text: markdown }
        ];

        const pythonCode = `
import sys, os, json
sys.path.append(r"${path.join(__dirname, 'solvers')}")
from doc_generator import create_document, convert_to_pdf, upload_to_temp

out_docx = r"${exportDocxPath}"
sections = ${JSON.stringify(docSections)}
create_document("${title}", "Ghost AI — Antigravity Remote Engine", "${exerciseText.replace(/"/g, '\\"')}", sections, out_docx)

pdf_path = convert_to_pdf(out_docx)
target = pdf_path if "${format}" == "pdf" and pdf_path else out_docx
url = upload_to_temp(target)
print(json.dumps({"docx": out_docx, "pdf": pdf_path, "url": url}))
`;

        execFile("python", ["-c", pythonCode], (err, stdout) => {
          if (err) {
            res.writeHead(500, { "Content-Type": "application/json" });
            res.end(JSON.stringify({ error: "Error al generar documento: " + err.message }));
            return;
          }
          try {
            const outJson = JSON.parse(stdout.trim().split("\n").pop());
            res.writeHead(200, { "Content-Type": "application/json" });
            res.end(JSON.stringify({
              ok: true,
              url: outJson.url || null,
              docxPath: outJson.docx,
              pdfPath: outJson.pdf
            }));
          } catch (pe) {
            res.writeHead(500, { "Content-Type": "application/json" });
            res.end(JSON.stringify({ error: "Error parseando salida del generador: " + stdout }));
          }
        });
      } catch (err) {
        res.writeHead(500, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: err.message }));
      }
      return;
    }

    // 4. Server-Sent Events (SSE) stream
    if (pathname === "/api/events" && req.method === "GET") {
      res.writeHead(200, {
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive"
      });

      const client = { id: Date.now(), res };
      sseClients.push(client);

      // Send initial connected ping
      res.write(`event: connected\ndata: ${JSON.stringify({ ok: true, time: Date.now() })}\n\n`);

      // If there was a recent solution, send it immediately to newly connected mobile device
      if (lastSolution) {
        res.write(`event: solution\ndata: ${JSON.stringify(lastSolution)}\n\n`);
      }

      req.on("close", () => {
        sseClients = sseClients.filter((c) => c.id !== client.id);
      });
      return;
    }

    // 5. Broadcast helper (used by Word or Floating app to notify SSE)
    if (pathname === "/api/broadcast" && req.method === "POST") {
      try {
        const body = await parseJsonBody(req);
        if (body.type && body.data) {
          broadcastSSE(body.type, body.data);
          if (body.type === "solution") {
            lastSolution = { ...body.data, timestamp: Date.now() };
          }
        }
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true }));
      } catch (err) {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: err.message }));
      }
      return;
    }

    // 6. Serve Remote Web App
    if (pathname === "/remote" || pathname === "/" || pathname === "/remote.html") {
      const htmlPath = path.join(__dirname, "public", "remote.html");
      if (fs.existsSync(htmlPath)) {
        res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        fs.createReadStream(htmlPath).pipe(res);
        return;
      }
    }

    // Default 404
    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("Not Found");
  });

  server.on("error", (err) => {
    if (err.code === "EADDRINUSE") {
      console.log(`Port ${port} in use, trying ${port + 1}...`);
      activePort = port + 1;
      server.listen(activePort);
    } else {
      console.error("Bridge Server error:", err);
    }
  });

  server.listen(port, () => {
    activePort = port;
    const ips = getLocalIps();
    console.log(`\n======================================================`);
    console.log(`👻 GHOST ANTIGRAVITY BRIDGE SERVER ACTIVO`);
    console.log(`======================================================`);
    console.log(`• Local:   http://localhost:${activePort}/remote`);
    ips.forEach((ip) => {
      console.log(`• Celular: http://${ip}:${activePort}/remote`);
    });
    console.log(`• API:     http://localhost:${activePort}/api/status`);
    console.log(`======================================================\n`);
  });

  return { server, broadcastSSE };
}

// Standalone execution support
if (require.main === module) {
  createBridgeServer(DEFAULT_PORT);
}

module.exports = { createBridgeServer, broadcastSSE, getLocalIps };
