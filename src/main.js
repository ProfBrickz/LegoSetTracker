// Imports
import ejs from "ejs";
import { app, BrowserWindow, nativeTheme } from "electron";
import fs from "fs";
import path from "path";
import { IMAGES_PATH, IS_DEV_MODE, LAYOUTS_PATH, MINIFIG_IMAGES_PATH, PAGES_PATH, PIECE_IMAGES_PATH, PRELOAD_FILE, SET_IMAGES_PATH } from "./constants.js";
import { LegoColors, LegoPieces, LegoSets, LegoSetThemes } from "./controllers.js";
import { ipcMain } from "./ipcMain.js";
import { LegoSet, LegoSetPiece } from "./models.js";
import WebScrapper from "./webScrapper.js";


// Variables
/** @type {BrowserWindow | null} */
let mainWindow;
let colors = new LegoColors();
let webScrapper = new WebScrapper(colors);
export { webScrapper };
let legoPieces = new LegoPieces();
let legoSets = new LegoSets();
LegoSet.legoPieces = legoPieces;
let legoSetThemes = new LegoSetThemes();


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
	let noColorImagePath = path.join(PIECE_IMAGES_PATH, "0");
	if (!fs.existsSync(noColorImagePath)) fs.mkdirSync(noColorImagePath);
}

/**
 * Gets the rows for the Lego set search table
 *
 * @param {import("./types.js").LegoSetSearchResult[]} searchResults
 */
async function getSearchLegoSetsTableRows(searchResults) {
	/** @type {(import("./types.js").LegoSetSearchResult & {theme: string, image: string})[]} */
	let tableRows = [];

	for (let searchResult of searchResults) {
		let themeIds = searchResult.themeId.split(".");
		let themes = [];

		for (let themeId of themeIds) {
			themes.push(legoSetThemes.getByBricklinkId(themeId));
		}

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

	return tableRows;
}

/**
 * Gets the rows for the main Lego sets table
 */
function getLegoSetsTableRows() {
	/**
	 * @type {import("./types.js").LegoSetsTableRow[]}
	 */
	let tableRows = [];

	for (let legoSet of legoSets.values()) {
		let image = fs.readFileSync(legoSet.getImagePath()).toString("base64") || "";

		tableRows.push({
			image: `data:image/jpg;base64,${image}`,
			databaseId: /** @type {number} */ (legoSet.databaseId),
			name: legoSet.name,
			setNumber: legoSet.setNumber,
			theme: legoSet.theme,
			releaseYear: legoSet.releaseYear,
			pieceCount: legoSet.pieceCount,
			minifigCount: legoSet.minifigCount,
			legoSetCount: legoSet.legoSetCount
		});
	}

	return tableRows;
}

/**
 * Gets the table rows for a Lego set
 *
 * @param {LegoSet} legoSet
 */
function getLegoSetTableRows(legoSet) {
	/**
	 * @type {import("./types.js").LegoSetTableRow[]}
	 */
	let tableRows = [];

	for (let legoPiece of legoSet.normalPieces.values()) {
		let image = fs.readFileSync(legoPiece.getImagePath()).toString("base64") || "";

		tableRows.push({
			image: `data:image/jpg;base64,${image}`,
			pieceId: /** @type {number} */ (legoPiece.databaseId),
			amountFound: legoPiece.amountFound,
			amountLeft: legoPiece.amountNeeded * legoSet.legoSetCount - legoPiece.amountFound,
			amountNeeded: legoPiece.amountNeeded * legoSet.legoSetCount,
			bricklinkName: legoPiece.bricklinkName,
			bricklinkId: legoPiece.bricklinkId,
			color: legoPiece.color,
			bricklinkCategory: legoPiece.bricklinkCategory
		});
	}

	return tableRows;
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

// Setup
await colors.init();
initializeFolders();
legoSetThemes.init();


// Event listeners
app.whenReady().then(async () => {
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
ipcMain.handle("loadPage", (event, page, { pageParams = {}, params = {} }) => {
	if (page === "add-lego-set") {
		pageParams.legoSetThemes = legoSetThemes;
	} else if (page == "lego-sets") {
		params.tableRows = getLegoSetsTableRows();
	} else if (page === "settings") {
		setTimeout(() => {
			mainWindow?.webContents.send("themeChange", nativeTheme.themeSource);
		}, 10);
	} else if (page == "lego-set") {
		let databaseId = /** @type {number} */ (params.databaseId);
		delete params.databaseId;

		let legoSet =/** @type {LegoSet} */(legoSets.get(databaseId));

		pageParams.legoSetName = legoSet.name;
		params.legoSetId = legoSet.databaseId;
		params.tableRows = getLegoSetTableRows(legoSet);
	}

	let html = renderPage(page, pageParams);

	return { page, html, params };
});

ipcMain.on("setTheme", async (event, theme) => {
	if (!mainWindow) return;

	nativeTheme.themeSource = theme;

	mainWindow.webContents.send("themeChange", theme);
});

ipcMain.handle("searchLegoSets", async (event, searchQuery, { themeId, startYear, endYear }) => {
	let searchResults = await webScrapper.searchLegoSets(searchQuery, { themeId, startYear, endYear });

	for (let searchResult of searchResults) {
		await webScrapper.downloadLegoSetImage(searchResult.setNumber);
	}

	return await getSearchLegoSetsTableRows(searchResults);
});

ipcMain.on("addLegoSet", async (event, setNumber) => {
	if (!mainWindow) return;

	let legoSetInfo;
	try {
		legoSetInfo = await webScrapper.getLegoSetInfo(setNumber);
	} catch (error) {
		return false;
	}

	let legoSet = legoSets.getBySetNumber(setNumber);

	// Increment set count if it already exists
	if (legoSet) {
		legoSet.legoSetCount++;
		return true;
	}

	// Get the pieces for this set
	let legoSetPieces = await webScrapper.getLegoSetPieces(setNumber);

	let theme = legoSetThemes.getByBricklinkId(legoSetInfo.themeId);
	if (!theme) return false;

	// Add Lego set
	legoSet = new LegoSet(
		legoSets.size,
		legoSetInfo.setNumber,
		legoSetInfo.name,
		theme,
		legoSetInfo.releaseYear,
		legoSetInfo.pieceCount,
		legoSetInfo.minifigCount
	);

	legoSet.addNormalPieces(legoSetPieces.normalPieces);
	legoSet.addMinifigs(legoSetPieces.minifigs);
	legoSet.addExtraPieces(legoSetPieces.extraPieces);
	legoSet.addCounterpartPieces(legoSetPieces.counterparts);
	legoSets.set(legoSets.size, legoSet);

	webScrapper.downloadLegoSetImages(legoSet);

	return true;
});

ipcMain.handle("deleteLegoSet", (event, databaseId) => {
	legoSets.delete(databaseId);
});

ipcMain.on("changeLegoSetCount", (event, legoSetId, setCount) => {
	let legoSet = /** @type {LegoSet} */ (legoSets.get(legoSetId));

	legoSet.legoSetCount = setCount;
});

ipcMain.on("changeAmountFound", (event, legoSetId, pieceId, amountFound) => {
	let legoSet = /** @type {LegoSet} */ (legoSets.get(legoSetId));
	let legoPiece = /** @type {LegoSetPiece} */ (legoSet?.normalPieces.get(pieceId));

	legoPiece.amountFound = amountFound;
});
