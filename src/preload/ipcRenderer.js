const electronIpcRenderer = require("electron").ipcRenderer;

/** @type {import("../types/electronIPC.d.ts").TypedIpcRenderer} */
export const ipcRenderer = electronIpcRenderer;
