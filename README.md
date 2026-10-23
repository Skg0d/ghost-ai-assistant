# 👻 Ghost AI — Study & Exam Assistant

Asistente flotante, sigiloso y superpuesto con soporte multi-documento impulsado por **Google Gemini 3.5 Flash**.

Diseñado para estudiar, repasar y responder consultas con el más alto rigor académico sin interferir con tus ventanas de navegación.

---

## ✨ Características Principales

* 🛡️ **Modo No-Foco (100% Sigiloso)**: La ventana utiliza flags nativos de Windows (`WS_EX_NOACTIVATE`). Podés hacer click, scrollear o consultar sin que tu navegador pierda el foco ni dispare eventos `blur` o `visibilitychange`.
* ✨ **Preguntar lo Seleccionado**: Resaltás cualquier texto en tu navegador con el mouse y con un solo click la IA lo consulta sin que toques el teclado (`Ctrl+C`).
* 📚 **Multi-Documento (RAG)**: Arrastrá uno o varios PDFs, apuntes o libros enteros directamente a la ventana.
* 🎯 **Rúbrica Académica Universitaria**: Las respuestas están calibradas para cumplir con los criterios de evaluación de ingeniería y sistemas:
  * Dominio conceptual y terminología formal.
  * Análisis estructurado y bien fundamentado.
  * Ejemplos prácticos pertinentes.
  * Justificación técnica de las decisiones.
  * Síntesis y cruce entre múltiples documentos.
* 💬 **Memoria Conversacional**: Recuerda las preguntas anteriores para responder repreguntas de seguimiento (*"¿según quién?"*, *"¿por qué?"*, *"¿cuál es la opción correcta?"*).
* 👁️ **Control de Opacidad**: Ajustá la transparencia de toda la ventana del 15% al 100% en tiempo real.

---

## 🚀 Cómo Usarlo

### Opción 1: Ejecutable Portable (Recomendado para compañeros sin Node.js)
1. Descargá el archivo **`GhostAI-Para-Companeros.zip`** desde la sección [Releases](../../releases).
2. Descomprimilo en cualquier carpeta.
3. Hacé doble click en **`GhostAI.exe`** y listo.

---

### Opción 2: Desde el código fuente (Con Node.js)
```bash
# 1. Clonar el repositorio
git clone <URL_DE_ESTE_REPOSITORIO>
cd ghost-ai

# 2. Instalar dependencias
npm install

# 3. Iniciar la app
npm start
```

---

## 🔑 Cómo obtener tu API Key gratuita de Gemini

Cada usuario debe configurar su propia API Key (es **100% gratuita**):

1. Ingresá a [Google AI Studio](https://aistudio.google.com/apikey) con tu cuenta de Google.
2. Hacé click en el botón azul **"Create API key"**.
3. Copiá la clave generada.
4. En Ghost AI, tocá el botón **`⚙️ API Key`** arriba a la derecha, pegala y dale a **Guardar**.
*(La clave queda guardada únicamente en tu máquina local).*

---

## ⌨️ Atajos de Teclado Globales

| Atajo | Acción |
|---|---|
| `Ctrl + Shift + H` | Ocultar / Mostrar la ventana |
| `Ctrl + Shift + A` | Preguntar lo que tengas copiado |
| `Ctrl + Shift + Space` | Activar / Bloquear modo teclado |
