import { contextBridge, ipcRenderer } from "electron";

contextBridge.exposeInMainWorld("hailCanva", {
  pickPptx: () => ipcRenderer.invoke("file:pickPptx"),
  listFonts: () => ipcRenderer.invoke("font:list"),
  savePptx: (name: string, bytes: Uint8Array) => ipcRenderer.invoke("file:savePptx", name, bytes),
  canvaStatus: () => ipcRenderer.invoke("canva:status"),
  connectCanva: () => ipcRenderer.invoke("canva:connect"),
  downloadCanvaPptx: (link: string) => ipcRenderer.invoke("canva:downloadPptx", link),
  inspectPublicCanva: (link: string) => ipcRenderer.invoke("canva:inspectPublic", link),
  openCanvaConsent: (link: string) => ipcRenderer.invoke("canva:openConsent", link),
  analyzePublicCanva: (link: string) => ipcRenderer.invoke("canva:analyzePublic", link),
});
