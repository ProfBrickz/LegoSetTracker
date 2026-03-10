// Imports
const { contextBridge, ipcRenderer } = require("electron");


// Context Bridge
/** @type {import("../renderer/types/electronAPI.js").ElectronAPI} */
let electronAPI = {
	loadPage: (page, pageParams = {}, params = {}) => {
		ipcRenderer.send("loadPage", page, pageParams, params);
	},
	setTheme: (theme) => {
		ipcRenderer.invoke("setTheme", theme);
	},
	searchLegoSets: () => {
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

		ipcRenderer.invoke("searchLegoSets", searchQuery, { themeId, startYear, endYear });
	},
	onSearchResults: (callback) => {
		ipcRenderer.on("searchResults", callback);
	},
	addSet: (setNumber) => {
		ipcRenderer.invoke("addSet", setNumber);
	}
};
contextBridge.exposeInMainWorld("electronAPI", electronAPI);

// IPC Listeners
ipcRenderer.on("pageLoad", (event, page, html, params) => {
	const mainElement = document.getElementById("main");
	if (!mainElement) return;

	mainElement.innerHTML = html;

	// Update active button
	const navButtons = document.querySelectorAll(".nav-link");
	for (let navButton of navButtons) {
		navButton.classList.remove("active");
	}

	// Add active class to current page button
	const currentButton = document.querySelector(`.nav-link[onclick*="${page}"]`);
	if (currentButton) {
		currentButton.classList.add("active");
	}

	// Notify that page has changed so icons can be recreated
	document.dispatchEvent(new CustomEvent("pageLoad", {
		detail: {
			page,
			params
		}
	}));
});

ipcRenderer.on("themeChange", (event, theme) => {
	const themeButtons = document.querySelectorAll(".theme-buttons button");
	for (let themeButton of themeButtons) {
		themeButton.classList.remove("active");
	}

	const currentButton = document.querySelector(`.theme-buttons button#theme-${theme}`);
	if (currentButton) {
		currentButton.classList.add("active");
	}
});
