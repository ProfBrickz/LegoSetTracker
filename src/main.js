// Imports
import ejs from "ejs";
import { app, BrowserWindow, ipcMain, nativeTheme, screen } from "electron";
import fs from "fs";
import path from "path";
import { IS_DEV_MODE, LAYOUTS_PATH, PAGES_PATH, SRC_DIRECTORY } from "./constants.js";
import { LegoColor } from "./models.js";
import WebScrapper from "./webScrapper.js";


// Variables
/** @type {BrowserWindow | null} */
let mainWindow;
/** @type {LegoColor[]} */
let colors = [];
let webScrapper = new WebScrapper(colors);


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
	let pageEJS = fs.readFileSync(pagePath, "utf-8");

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
	let { width, height } = screen.getPrimaryDisplay().workAreaSize;

	mainWindow = new BrowserWindow({
		width,
		height,
		webPreferences: {
			contextIsolation: true,
			devTools: IS_DEV_MODE,
			preload: path.join(SRC_DIRECTORY, "preload.js")
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
	(event, page, params) => {
		if (!mainWindow) return;

		let html = loadPage(page, params);

		mainWindow.webContents.send("pageLoaded", page, html);

		if (page == "settings") {
			mainWindow.webContents.send("themeChange", nativeTheme.themeSource);
		}
	}
);

ipcMain.handle("setTheme",
	/**
	 * @param {Electron.IpcMainInvokeEvent} event
	 * @param {"light" | "dark" | "system"} theme The theme to set.
	 * @returns {void}
	 */
	(event, theme) => {
		if (!mainWindow) return;

		nativeTheme.themeSource = theme;

		mainWindow.webContents.send("themeChange", theme);
	}
);

/**
	 * @param {string} searchQuery The search query for the search.
	 * @param {Object} [options]
	 * @param {string} [options.themeId] The ID of the theme to filter by.
	 * @param {string} [options.startYear] The start year for the search (inclusive).
	 * @param {string} [options.endYear] The end year for the search (inclusive).
	 */
ipcMain.handle("searchLegoSets", async (event, searchQuery, { themeId = "", startYear = "", endYear = "" } = {}) => {
	if (!mainWindow) return;

	// let searchResults = await webScrapper.searchLegoSets(searchQuery, { themeId, startYear, endYear });
	let searchResults = [
		{
			"setNumber": "75033-1",
			"name": "Star Destroyer",
			"themeId": "65.806.258"
		},
		{
			"setNumber": "8303-1",
			"name": "Demon Destroyer",
			"themeId": "179.571"
		},
		{
			"setNumber": "8002-1",
			"name": "Destroyer Droid",
			"themeId": "36.65.257"
		}
	];

	mainWindow.webContents.send("searchResults", searchResults);
});
