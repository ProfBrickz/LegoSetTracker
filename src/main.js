// Imports
import ejs from "ejs";
import { app, BrowserWindow, ipcMain, nativeTheme } from "electron";
import fs from "fs";
import path from "path";
import { IS_DEV_MODE, LAYOUTS_PATH, PAGES_PATH, PRELOAD_FILE } from "./constants.js";
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
	let pagePath = PAGES_PATH;
	if (process.env['ELECTRON_RENDERER_URL']) {
		pagePath = path.resolve("src/renderer/views/pages");
	}
	pagePath = path.join(pagePath, `${page}.ejs`);
	let pageEJS = fs.readFileSync(pagePath, "utf-8");

	let html = ejs.render(pageEJS, params);

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

	if (process.env['ELECTRON_RENDERER_URL']) {
		mainWindow.loadURL(new URL(`views/layouts/${layout}.html`, process.env['ELECTRON_RENDERER_URL']).href);
	} else {
		mainWindow.loadFile(path.join(LAYOUTS_PATH, `${layout}.html`));
	}
}

/**
 * Creates a new BrowserWindow instance and loads the main application view.
 */
function createWindow() {
	mainWindow = new BrowserWindow({
		show: false,
		webPreferences: {
			contextIsolation: true,
			devTools: IS_DEV_MODE,
			preload: PRELOAD_FILE
		}
	});

	mainWindow.maximize();
	mainWindow.show();

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

		let html = renderPage(page, params);

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
