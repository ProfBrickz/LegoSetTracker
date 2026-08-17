// Imports
import ejs from "ejs";
import { app, BrowserWindow, nativeTheme } from "electron";
import fs from "fs";
import path from "path";
import { DATA_PATH, IMAGES_PATH, IS_BUILT, LAYOUTS_PATH, MINIFIG_IMAGES_PATH, PAGES_PATH, PIECE_IMAGES_PATH, PRELOAD_FILE, SET_IMAGES_PATH } from "./constants.js";
import database from "./database/database.js";
import { LegoColors, LegoPieces, LegoSets, LegoSetThemes } from "./dataMaps.js";
import { ipcMain } from "./ipcMain.js";
import { CompoundLegoSetPiece, LegoSet, LegoSetPiece, StickeredLegoSetPiece } from "./models.js";
import WebScraper from "./webScraper.js";
/** @import {LegoSetPieceType, LegoSetSearchResult, LegoSetsTableRow, LegoSetTableRow, LegoSetStatus} from "./types.js" */


// Variables
/** @type {BrowserWindow | null} */
let mainWindow;
let colors = new LegoColors();
let webScraper = new WebScraper(colors);
await webScraper.init();
export { webScraper };
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
 *
 * @returns {string} The rendered HTML.
 */
function renderPage(page, params = {}) {
	let pagePath = PAGES_PATH;
	if (process.env["ELECTRON_RENDERER_URL"]) {
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
 *
 * @returns {void}
 */
function loadLayout(layout) {
	if (!mainWindow) return;

	if (process.env["ELECTRON_RENDERER_URL"]) {
		mainWindow.loadURL(new URL(`views/layouts/${layout}.html`, process.env["ELECTRON_RENDERER_URL"]).href);
	} else {
		mainWindow.loadFile(path.join(LAYOUTS_PATH, `${layout}.html`));
	}
}

/**
 * Initializes all models
 */
async function initializeModels() {
	await colors.init(webScraper);
	await legoSetThemes.init(webScraper);
	await legoPieces.init(colors);
	await legoSets.init(legoSetThemes, legoPieces);
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
		let colorImagePath = path.join(PIECE_IMAGES_PATH, color.brickLinkId.toString());

		if (!fs.existsSync(colorImagePath)) fs.mkdirSync(colorImagePath);
	}
	let noColorImagePath = path.join(PIECE_IMAGES_PATH, "0");
	if (!fs.existsSync(noColorImagePath)) fs.mkdirSync(noColorImagePath);
}

/**
 * Gets the rows for the Lego set search table
 *
 * @param {LegoSetSearchResult[]} searchResults
 */
async function getSearchLegoSetsTableRows(searchResults) {
	/** @type {(LegoSetSearchResult & {theme: string, image: string})[]} */
	let tableRows = [];

	for (let searchResult of searchResults) {
		let tableRow = {
			...searchResult,
			theme: "",
			image: ""
		};

		let theme = legoSetThemes.getByBricklinkId(searchResult.themeId);
		if (theme) tableRow.theme = theme.brickLinkName;

		let image = fs.readFileSync(LegoSet.getImagePath(searchResult.setNumber)).toString("base64") || "";
		tableRow.image = `data:image/png;base64,${image}`;

		tableRows.push(tableRow);
	}

	return tableRows;
}

/**
 * Gets the rows for the main Lego sets table
 */
function makeLegoSetsTableRows() {
	/** @type {LegoSetsTableRow[]} */
	let tableRows = [];

	for (let legoSet of legoSets.values()) {
		let image = fs.readFileSync(legoSet.getImagePath()).toString("base64") || "";

		tableRows.push({
			image: `data:image/png;base64,${image}`,
			databaseId: /** @type {number} */ (legoSet.databaseId),
			name: legoSet.name,
			setNumber: legoSet.setNumber,
			theme: legoSet.theme.brickLinkName,
			releaseYear: legoSet.releaseYear,
			pieceCount: legoSet.pieceCount,
			minifigCount: legoSet.minifigCount,
			legoSetCount: legoSet.legoSetCount,
			grayedOut: legoSet.status === "scraping"
		});
	}

	return tableRows;
}

/**
 * @param {LegoSetTableRow[]} tableRows
 * @param {number} databaseId
 *
 * @returns {LegoSetTableRow | null}
 */
function getLegoSetPieceTableRow(tableRows, databaseId) {
	for (let tableRow of tableRows) {
		if (tableRow.databaseId === databaseId) return tableRow;

		if (tableRow.subRows.length > 0) {
			let foundRow = getLegoSetPieceTableRow(tableRow.subRows, databaseId);
			if (foundRow) return foundRow;
		}
	}

	return null;
}

/**
 * @param {LegoSetTableRow[]} tableRows
 * @param {number} databaseId
 *
 * @returns {boolean}
 */
function removeLegoSetPieceTableRow(tableRows, databaseId) {
	for (let i = 0; i < tableRows.length; i++) {
		let tableRow = tableRows[i];

		if (tableRow.databaseId === databaseId) {
			tableRows.splice(i, 1);
			return true;
		}
		if (tableRow.subRows.length > 0) {
			let inSubRows = removeLegoSetPieceTableRow(tableRow.subRows, databaseId);
			if (inSubRows) return true;
		}
	}

	return false;
}

/**
 * @param {LegoSet} legoSet
 * @param {LegoSetPiece} legoSetPiece
 * @param {LegoSetPieceType} legoSetPieceType
 *
 * @returns {LegoSetTableRow}
 */
function makeLegoSetPieceTableRow(legoSet, legoSetPiece, legoSetPieceType) {
	let image = "";

	if (legoSetPieceType === "minifig") image = legoSetPiece.getMinifigImagePath();
	else image = legoSetPiece.getImagePath();

	image = fs.readFileSync(image).toString("base64") || "";

	let min = 0;

	if (legoSetPiece instanceof CompoundLegoSetPiece) min = legoSetPiece.getMaxStickeredPiecesFound();
	else min = legoSetPiece.getTotalStickeredPiecesFound();

	return {
		image: `data:image/png;base64,${image}`,
		rowId: /** @type {number} */ (legoSetPiece.databaseId),
		databaseId: /** @type {number} */ (legoSetPiece.databaseId),
		subRows: [],
		amountFound: legoSetPiece.amountFound,
		amountLeft: legoSetPiece.amountNeeded * legoSet.legoSetCount - legoSetPiece.amountFound,
		amountNeeded: legoSetPiece.amountNeeded * legoSet.legoSetCount,
		brickLinkName: legoSetPiece.brickLinkName,
		brickLinkId: legoSetPiece.brickLinkId,
		color: legoSetPiece.color,
		brickLinkCategory: legoSetPiece.brickLinkCategory,
		min,
		legoSetCount: legoSet.legoSetCount
	};
}

/**
 * Gets the table rows for a Lego set
 *
 * @param {LegoSet} legoSet
 */
function makeLegoSetTableRows(legoSet) {
	/** @type {LegoSetTableRow[]}	 */
	let tableRows = [];

	for (let legoSetPiece of legoSet.normalPieces.values()) {
		tableRows.push(makeLegoSetPieceTableRow(legoSet, legoSetPiece, "normal"));
	}

	for (let legoSetPiece of legoSet.minifigs.values()) {
		if (legoSetPiece instanceof CompoundLegoSetPiece) {
			let tableRow = makeLegoSetPieceTableRow(legoSet, legoSetPiece, "minifig");

			for (let componentLegoSetPiece of legoSetPiece.componentLegoSetPieces) {
				let componentRow = getLegoSetPieceTableRow(tableRows, componentLegoSetPiece.databaseId);

				if (componentRow === null) {
					tableRow.subRows.push(makeLegoSetPieceTableRow(legoSet, componentLegoSetPiece, "minifigPieces"));
				} else {
					tableRow.subRows.push(componentRow);
				}
			}

			tableRows.push(tableRow);
			continue;
		}

		tableRows.push(makeLegoSetPieceTableRow(legoSet, legoSetPiece, "minifig"));
	}

	for (let legoSetPiece of legoSet.counterpartPieces.values()) {
		if (legoSetPiece instanceof StickeredLegoSetPiece) {
			let baseRow = getLegoSetPieceTableRow(tableRows, legoSetPiece.baseLegoSetPiece.databaseId);

			if (baseRow !== null) {
				baseRow.subRows.push(makeLegoSetPieceTableRow(legoSet, legoSetPiece, "counterpart"));
				continue;
			}
		} else if (legoSetPiece instanceof CompoundLegoSetPiece) {
			let tableRow = makeLegoSetPieceTableRow(legoSet, legoSetPiece, "counterpart");

			for (let componentLegoSetPiece of legoSetPiece.componentLegoSetPieces) {
				let componentRow = getLegoSetPieceTableRow(tableRows, componentLegoSetPiece.databaseId);

				if (componentRow !== null) {
					tableRow.subRows.push(componentRow);
					removeLegoSetPieceTableRow(tableRows, componentRow.databaseId);
				}
			}

			tableRows.push(tableRow);
			continue;
		}

		tableRows.push(makeLegoSetPieceTableRow(legoSet, legoSetPiece, "counterpart"));
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
			devTools: !IS_BUILT,
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
if (!fs.existsSync(DATA_PATH)) fs.mkdirSync(DATA_PATH);
await database.init();
await initializeModels();
initializeFolders();


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

app.on("window-all-closed", async () => {
	// Quit when all windows are closed, except on macOS
	if (process.platform === "darwin") return;

	await webScraper.close();
	app.quit();
});


// IPC
ipcMain.handle("loadPage", (event, page, { pageParams = {}, params = {} }) => {
	if (page === "add-lego-set") {
		pageParams.legoSetThemes = legoSetThemes;
	} else if (page === "lego-sets") {
		params.tableRows = makeLegoSetsTableRows();
	} else if (page === "settings") {
		setTimeout(() => {
			mainWindow?.webContents.send("themeChange", nativeTheme.themeSource);
		}, 10);
	} else if (page === "lego-set") {
		let databaseId = /** @type {number} */ (params.databaseId);
		delete params.databaseId;

		let legoSet =/** @type {LegoSet} */(legoSets.get(databaseId));

		pageParams.legoSetName = legoSet.name;
		pageParams.legoSetCount = legoSet.legoSetCount;
		params.legoSetId = legoSet.databaseId;
		params.tableRows = makeLegoSetTableRows(legoSet);
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
	let searchResults = await webScraper.searchLegoSets(searchQuery, { themeId, startYear, endYear });

	for (let searchResult of searchResults) {
		await webScraper.downloadLegoSetImage(searchResult.setNumber);
	}

	return await getSearchLegoSetsTableRows(searchResults);
});

ipcMain.on("addLegoSet", async (event, setNumber) => {
	if (!mainWindow) return;

	let legoSet = legoSets.getBySetNumber(setNumber);

	// Increment set count if it already exists
	if (legoSet) {
		legoSet.legoSetCount++;
		legoSet.save();
		return true;
	}

	let legoSetInfo;
	try {
		legoSetInfo = await webScraper.getLegoSetInfo(setNumber);
	} catch (error) {
		return false;
	}

	// Get the pieces for this set
	let legoSetPieces = await webScraper.getLegoSetPieces(setNumber);

	let theme = legoSetThemes.getByBricklinkId(legoSetInfo.themeId);
	if (!theme) return false;

	// Add Lego set
	legoSet = await legoSets.add(
		legoSetInfo.setNumber,
		legoSetInfo.name,
		theme,
		legoSetInfo.releaseYear,
		legoSetInfo.pieceCount,
		legoSetInfo.minifigCount
	);

	await legoSet.addNormalPieces(legoSetPieces.normalPieces);
	await legoSet.addMinifigs(legoSetPieces.minifigs);
	await legoSet.addExtraPieces(legoSetPieces.extraPieces);
	await legoSet.addCounterpartPieces(legoSetPieces.counterparts);
	legoSet.status = "done";
	legoSet.save();

	webScraper.downloadLegoSetImages(legoSet);

	return true;
});

ipcMain.handle("deleteLegoSet", (event, databaseId) => {
	legoSets.delete(databaseId);
});

ipcMain.on("changeLegoSetCount", (event, legoSetId, legoSetCount) => {
	let legoSet = /** @type {LegoSet} */ (legoSets.get(legoSetId));

	legoSet.legoSetCount = legoSetCount;
	legoSet.save();
});

ipcMain.handle("changeAmountFound", (event, legoSetId, pieceId, amountFound) => {
	/** @type {LegoSetPiece[]} */
	let result = [];

	let legoSet = /** @type {LegoSet} */ (legoSets.get(legoSetId));
	let legoSetPiece = /** @type {LegoSetPiece} */ (legoSet?.normalPieces.get(pieceId));
	if (legoSetPiece === undefined) legoSetPiece = /** @type {LegoSetPiece} */ (legoSet?.minifigs.get(pieceId));
	if (legoSetPiece === undefined) legoSetPiece = /** @type {LegoSetPiece} */ (legoSet?.extraPieces.get(pieceId));
	if (legoSetPiece === undefined) legoSetPiece = /** @type {LegoSetPiece} */ (legoSet?.counterpartPieces.get(pieceId));
	if (legoSetPiece === undefined) legoSetPiece = /** @type {LegoSetPiece} */ (legoSet?.minifigPieces.get(pieceId));

	legoSetPiece.syncAmountFound(amountFound, (legoSetPiece) => {
		legoSetPiece.save();
		result.push(legoSetPiece);
	});

	return result;
});
