// Imports
const electronIpcMain = require("electron").ipcMain;
/** @import {TypedIpcMain} from "./types/electronIPC.d.ts" */


/** @type {TypedIpcMain} */
export const ipcMain = electronIpcMain;
