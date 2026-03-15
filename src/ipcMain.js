const electronIpcMain = require("electron").ipcMain;

/** @type {import("./types/electronIPC.d.ts").TypedIpcMain} */
export const ipcMain = electronIpcMain;
