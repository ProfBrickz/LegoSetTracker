// Imports
import ejs from 'ejs';
import { app, BrowserWindow, ipcMain } from "electron";
import fs from 'fs';
import path from 'path';
import { DIR_NAME, IS_DEV_MODE } from "./constants.js";


// Constants
const VIEWS_PATH = path.join(DIR_NAME, "views");
const PAGES_PATH = path.join(VIEWS_PATH, "pages");
const LAYOUTS_PATH = path.join(VIEWS_PATH, "layouts");


// Constants
// Variables
/** @type {BrowserWindow | null} */
let mainWindow;


// Functions
/**
 * Renders an EJS template with the given parameters.
 *
 * @param {string} page The name of the EJS template file to render.
 * @param {Object} [params] The parameters to pass to the EJS template.
 * @returns {string} The rendered HTML.
 */
function renderPage(page, params = {}) {
	let pagePath = path.join(PAGES_PATH, `${page}.ejs`);
	let pageEJS = fs.readFileSync(pagePath, 'utf-8');

	let html = ejs.render(pageEJS, params);

	return html;
}


/**
 * Loads the specified page into the main window.
 *
 * @param {string} page The name of the page to load.
 * @param {Object} [params] The parameters to pass to the page.
 * @returns {string}
 */
function loadPage(page, params = {}) {
	if (!mainWindow) return "";

	let html = renderPage(page, params);
	return html;
}

/**
 * Loads a layout HTML file into the main window.
 *
 * @param {string} layout The name of the layout to load.
 * @returns {void}
 */
function loadLayout(layout) {
	if (!mainWindow) return;

	mainWindow.loadFile(path.join(LAYOUTS_PATH, `${layout}.html`));
}

/**
 * Creates a new BrowserWindow instance and loads the main application view.
 */
function createWindow() {
	mainWindow = new BrowserWindow({
		width: 960,
		height: 540,
		webPreferences: {
			contextIsolation: true,
			devTools: IS_DEV_MODE,
			preload: path.join(DIR_NAME, "preload.js")
		}
	});

	loadLayout("main");

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


// IPC
ipcMain.on("loadPage",
	/**
	 * @param {Electron.IpcMainEvent} event
	 * @param {Object} data
	 * @param {string} data.page The page to navigate to.
	 * @param {Object} data.params The parameters for the page.
	 * @returns {void}
	 */
	(event, { page, params }) => {
		if (!mainWindow) return;

		let html = loadPage(page, params);

		mainWindow.webContents.send("pageLoaded", { html });
	}
);
