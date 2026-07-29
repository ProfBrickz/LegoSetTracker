// Imports
const { contextBridge } = require("electron");
const electronIpcRenderer = require("electron").ipcRenderer;
/** @import {TypedIpcRenderer} from "../types/electronIPC.d.ts" */
/** @import {ElectronAPI} from "../renderer/types/electronAPI.d.ts" */


/** @type {TypedIpcRenderer} */
export const ipcRenderer = electronIpcRenderer;


// Context Bridge
/** @type {ElectronAPI} */
let electronAPI = {
	loadPage: async (page, { pageParams = {}, params = {} } = {}) => {
		// call loadPage on server side
		let { page: newPage, html, params: newParams } = await ipcRenderer.invoke("loadPage", page, { pageParams, params });

		const mainElement = document.getElementById("main");
		if (!mainElement) return;

		// Update page content
		mainElement.innerHTML = html;

		// Update active button
		const navButtons = document.querySelectorAll(".nav-link");
		for (let navButton of navButtons) {
			navButton.classList.remove("active");
		}

		// Add active class to current page button
		const currentButton = document.querySelector(`.nav-link[onclick*="${newPage}"]`);
		if (currentButton) {
			currentButton.classList.add("active");
		}

		// Notify that page has changed, so the js for each page can detect it
		document.dispatchEvent(new CustomEvent("pageLoad", {
			detail: {
				page: newPage,
				params: newParams
			}
		}));
	},
	setTheme: async (theme) => {
		ipcRenderer.send("setTheme", theme);
	},
	searchLegoSets: async (callback) => {
		// Get form values
		let searchQueryElement = /** @type {HTMLInputElement} */ (document.getElementById("search-query"));
		let searchQuery = searchQueryElement.value;
		let themeElement =/** @type {HTMLSelectElement} */ (document.getElementById("theme"));
		let themeId = themeElement.value;
		let startYearElement = /** @type {HTMLInputElement} */ (document.getElementById("start-year"));
		let startYear;
		if (!Number.isNaN(startYear)) startYear = startYearElement.valueAsNumber;
		let endYearElement = /** @type {HTMLInputElement} */ (document.getElementById("end-year"));
		let endYear;
		if (!Number.isNaN(endYear)) endYear = endYearElement.valueAsNumber;

		let tableRows = await ipcRenderer.invoke("searchLegoSets", searchQuery, { themeId, startYear, endYear });
		callback(tableRows);
	},
	addLegoSet: (setNumber) => {
		ipcRenderer.send("addLegoSet", setNumber);
	},
	deleteLegoSet: async (databaseId) => {
		await ipcRenderer.invoke("deleteLegoSet", databaseId);
	},
	changeLegoSetCount: (legoSetId, setCount) => {
		ipcRenderer.send("changeLegoSetCount", legoSetId, setCount);
	},
	changeAmountFound: async (legoSetId, pieceId, amountFound) => {
		return await ipcRenderer.invoke("changeAmountFound", legoSetId, pieceId, amountFound);
	}
};
contextBridge.exposeInMainWorld("electronAPI", electronAPI);


// IPC Listeners
ipcRenderer.on("themeChange", (event, theme) => {
	const themeButtons = /** @type {NodeListOf<HTMLButtonElement>} */ (document.querySelectorAll("#theme-buttons button"));
	for (let themeButton of themeButtons) {
		themeButton.classList.remove("active");
	}

	const currentButton = /** @type {HTMLButtonElement | null} */ (document.querySelector(`#theme-buttons button#theme-${theme}`));
	if (currentButton) {
		currentButton.classList.add("active");
	}
});
