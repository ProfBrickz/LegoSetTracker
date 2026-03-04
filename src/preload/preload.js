// Imports
const { contextBridge, ipcRenderer } = require("electron");


// Context Bridge
contextBridge.exposeInMainWorld("electronAPI", {
	/**
	 * @param {string} page The page to navigate to.
	 * @param {Object} pageParams The parameters for the page.
	 * @param {Object} params The parameters for the page.
	 * @returns {void}
	 */
	loadPage: (page, pageParams = {}, params = {}) => {
		ipcRenderer.send("loadPage", page, pageParams);
	},
	/**
	 * @param {import("../types.js").Theme} theme The theme to set.
	 * @returns {void}
	 */
	setTheme: (theme) => {
		ipcRenderer.invoke("setTheme", theme);
	},
	/**
	 * @returns {void}
	 */
	searchLegoSets: () => {
		let searchQueryElement = /** @type {HTMLInputElement} */ (document.getElementById("search-query"));
		let searchQuery = searchQueryElement.value;
		let startYearElement = /** @type {HTMLInputElement} */ (document.getElementById("start-year"));
		let startYear = startYearElement.value;
		let endYearElement = /** @type {HTMLInputElement} */ (document.getElementById("end-year"));
		let endYear = endYearElement.value;

		ipcRenderer.invoke("searchLegoSets", searchQuery, { startYear, endYear });
	}
});

// IPC Listeners
ipcRenderer.on("pageLoaded", (event, page, html) => {
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
	document.dispatchEvent(new CustomEvent("pageChanged"));
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

ipcRenderer.on("searchResults", (event, searchResults) => {
	console.log(searchResults);
});
