// Imports
import { app, BrowserWindow } from "electron";
import { IS_DEV_MODE } from "./constants.js";


// Variables
/**
 * @type {BrowserWindow | null}
 */
let mainWindow;


// Functions
/**
 * Creates a new BrowserWindow instance and loads the main application view.
 */
function createWindow() {
	mainWindow = new BrowserWindow({
		width: 960,
		height: 540,
		webPreferences: {
			nodeIntegration: true,
			contextIsolation: false,
			devTools: IS_DEV_MODE,
		}
	});

	mainWindow.loadFile("src/views/layout.html");

	mainWindow.on("closed", () => {
		mainWindow = null;
	});
}


// Event listeners
app.whenReady().then(() => {
	createWindow();

	app.on("activate", () => {
		// On macOS it's common to re-create a window in the app when the
		// dock icon is clicked and there are no other windows open
		if (mainWindow === null || BrowserWindow.getAllWindows().length === 0) {
			createWindow();
		}
	});
});

app.on("window-all-closed", () => {
	// Quit when all windows are closed, except on macOS
	if (process.platform !== "darwin") {
		app.quit();
	}
});
