// Imports
import ejs from "ejs";
import { app, BrowserWindow, ipcMain, nativeTheme } from "electron";
import fs from "fs";
import path from "path";
import { IMAGES_PATH, IS_DEV_MODE, LAYOUTS_PATH, MINIFIG_IMAGES_PATH, PAGES_PATH, PIECE_IMAGES_PATH, PRELOAD_FILE, SET_IMAGES_PATH } from "./constants.js";
import { LegoColor, LegoSet } from "./models.js";
import WebScrapper from "./webScrapper.js";


// Variables
/** @type {BrowserWindow | null} */
let mainWindow;
/** @type {LegoColor[]} */
let colors = [];
/** @type {Map<string, string>} */
let legoSetThemes = new Map();
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
 * Initializes all necessary folders.
 */
function initializeFolders() {
	if (!fs.existsSync(IMAGES_PATH)) fs.mkdirSync(IMAGES_PATH);
	if (!fs.existsSync(SET_IMAGES_PATH)) fs.mkdirSync(SET_IMAGES_PATH);
	if (!fs.existsSync(PIECE_IMAGES_PATH)) fs.mkdirSync(PIECE_IMAGES_PATH);
	if (!fs.existsSync(MINIFIG_IMAGES_PATH)) fs.mkdirSync(MINIFIG_IMAGES_PATH);
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
app.whenReady().then(async () => {
	initializeFolders();
	legoSetThemes = await webScrapper.getLegoSetThemes();

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
	 * @param {string} page The page to navigate to.
	 * @param {Object} pageParams The parameters for the page ejs.
	 * @param {Object} params The parameters for the page js.
	 * @returns {void}
	 */
	(event, page, pageParams, params) => {
		if (!mainWindow) return;

		let html = renderPage(page, pageParams);

		mainWindow.webContents.send("pageLoad", page, html, params);

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
	/** @type {import("./types.js").LegoSetSearchResult[]} */
	let searchResults = [
		{
			setNumber: "75033-1",
			name: "Star Destroyer",
			themeId: "65.806.258"
		},
		{
			setNumber: "8303-1",
			name: "Demon Destroyer",
			themeId: "179.571"
		},
		{
			setNumber: "8002-1",
			name: "Destroyer Droid",
			themeId: "36.65.257"
		}
	];
	/** @type {(import("./types.js").LegoSetSearchResult & {theme: string, image: string})[]} */
	let tableRows = [];

	for (let searchResult of searchResults) {
		let themeIds = searchResult.themeId.split(".");
		let themes = [];

		for (let themeId of themeIds) {
			themes.push(legoSetThemes.get(themeId));
		}

		await webScrapper.downloadLegoSetImage(searchResult.setNumber);

		let tableRow = {
			...searchResult,
			theme: "",
			image: ""
		};
		tableRow.theme = themes.join(": ");
		let image = fs.readFileSync(LegoSet.getImagePath(searchResult.setNumber)).toString("base64") || "";
		tableRow.image = `data:image/jpg;base64,${image}`;

		tableRows.push(tableRow);
	}

	mainWindow.webContents.send("searchResults", tableRows);
});
