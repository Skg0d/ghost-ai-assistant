const { contextBridge, ipcRenderer, webUtils } = require("electron");
const { marked } = require("marked");

// Configure marked options for clean rendering
marked.setOptions({
  gfm: true,
  breaks: true,
});

contextBridge.exposeInMainWorld("ghostAPI", {
  renderMarkdown: (text) => {
    try {
      return marked.parse(text || "");
    } catch {
      return text || "";
    }
  },
  sendMessage: (data) => ipcRenderer.invoke("ai-request", data),
  setOpacity: (val) => ipcRenderer.invoke("set-opacity", val),
  getClipboard: () => ipcRenderer.invoke("get-clipboard"),
  getSelectedText: () => ipcRenderer.invoke("get-selected-text"),
  selectDocument: () => ipcRenderer.invoke("select-document"),
  loadDocumentPath: (filePath) => ipcRenderer.invoke("load-document-path", filePath),
  loadDocumentBuffer: (name, buffer) => ipcRenderer.invoke("load-document-buffer", { name, buffer }),
  getPathForFile: (file) => {
    try {
      if (webUtils && typeof webUtils.getPathForFile === "function") {
        return webUtils.getPathForFile(file);
      }
    } catch {}
    return file.path || "";
  },
  setFocusMode: (state) => ipcRenderer.invoke("set-focus-mode", state),
  getRemoteInfo: () => ipcRenderer.invoke("get-remote-info"),
  onAskClipboard: (cb) => ipcRenderer.on("ask-clipboard", cb),
  onFocusStateChanged: (cb) => ipcRenderer.on("focus-state-changed", (_e, state) => cb(state)),
  getProjects: () => ipcRenderer.invoke("get-projects"),
  getConversationHistory: (cid) => ipcRenderer.invoke("get-conversation-history", cid),
  createConversation: (data) => ipcRenderer.invoke("create-conversation", data),
  selectProjectFolder: () => ipcRenderer.invoke("select-project-folder"),
});
