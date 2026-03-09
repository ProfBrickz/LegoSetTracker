// Imports
import ejs from "ejs";
import { app, BrowserWindow, ipcMain, nativeTheme } from "electron";
import fs from "fs";
import path from "path";
import { IMAGES_PATH, IS_DEV_MODE, LAYOUTS_PATH, MINIFIG_IMAGES_PATH, PAGES_PATH, PIECE_IMAGES_PATH, PRELOAD_FILE, SET_IMAGES_PATH } from "./constants.js";
import { LegoColors, LegoPieces, LegoSets } from "./controllers.js";
import { LegoSet } from "./models.js";
import WebScrapper from "./webScrapper.js";


// Variables
/** @type {BrowserWindow | null} */
let mainWindow;
let colors = new LegoColors();
let legoPieces = new LegoPieces();
let legoSets = new LegoSets();
LegoSet.legoPieces = legoPieces;
/** @type {Map<string, string>} */
let legoSetThemes = new Map();
let webScrapper = new WebScrapper(colors);


// Functions
/**
 * Renders an EJS template with the given parameters.
 *
 * @param {string} page The name of the EJS template file to render.
 * @param {Record<string, unknown>} [params] The parameters to pass to the EJS template.
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

	for (let color of colors.values()) {
		let colorImagePath = path.join(PIECE_IMAGES_PATH, color.bricklinkId.toString());

		if (!fs.existsSync(colorImagePath)) fs.mkdirSync(colorImagePath);
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
app.whenReady().then(async () => {
	legoSetThemes = await webScrapper.getLegoSetThemes();
	colors = await webScrapper.getColors();
	webScrapper.setColors(colors);

	initializeFolders();

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
	 * @param {Record<string, unknown>} pageParams The parameters for the page ejs.
	 * @param {Record<string, unknown>} params The parameters for the page js.
	 * @returns {void}
	 */
	(event, page, pageParams, params) => {
		if (!mainWindow) return;

		if (page === "add-set") {
			pageParams.legoSetThemes = legoSetThemes;
		}

		let html = renderPage(page, pageParams);

		mainWindow.webContents.send("pageLoad", page, html, params);

		if (page === "settings") {
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

ipcMain.handle("searchLegoSets",
	/**
	 * @param {Electron.IpcMainInvokeEvent} event
	 * @param {string} searchQuery The search query for the search.
	 * @param {Object} [options]
	 * @param {string} [options.themeId] The ID of the theme to filter by.
	 * @param {string} [options.startYear] The start year for the search (inclusive).
	 * @param {string} [options.endYear] The end year for the search (inclusive).
	 */
	async (event, searchQuery, { themeId = "", startYear = "", endYear = "" } = {}) => {
		if (!mainWindow) return;

		let searchResults = await webScrapper.searchLegoSets(searchQuery, { themeId, startYear, endYear });
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

ipcMain.handle("addSet",
	/**
	 * @param {Electron.IpcMainInvokeEvent} event
	 * @param {string} setNumber
	 */
	async (event, setNumber) => {
		if (!mainWindow) return;

		let legoSetInfo;
		try {
			legoSetInfo = await webScrapper.getLegoSetInfo(setNumber);
		} catch (error) {
			return false;
		}

		let legoSetPieces = await webScrapper.getLegoSetPieces(setNumber);

		let legoSet = new LegoSet(
			legoSets.size,
			legoSetInfo.setNumber,
			legoSetInfo.name,
			legoSetInfo.theme,
			legoSetInfo.releaseYear,
			legoSetInfo.pieceCount,
			legoSetInfo.minifigCount
		);

		legoSet.addNormalPieces(legoSetPieces.normalPieces);
		legoSet.addMinifigs(legoSetPieces.minifigs);
		legoSet.addExtraPieces(legoSetPieces.extraPieces);
		legoSet.addCounterpartPieces(legoSetPieces.counterparts);

		await webScrapper.downloadLegoSetImages(legoSet);

		mainWindow.webContents.send("addSet");
	});
