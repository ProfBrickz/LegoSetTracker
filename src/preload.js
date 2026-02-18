// Imports
const { contextBridge, ipcRenderer } = require("electron");

// Cache for rendered pages to improve performance
const pageCache = new Map();
let currentPage = null;

contextBridge.exposeInMainWorld("electronAPI", {
	/**
	 * @param {string} page The page to navigate to.
	 * @param {Object} params The parameters for the page.
	 * @returns {void}
	 */
	loadPage: (page, params = {}) => {
		ipcRenderer.send("loadPage", { page, params });
	},
	/**
	 * @param {"light" | "dark" | "system"} theme The theme to set.
	 * @returns {void}
	 */
	setTheme: (theme) => {
		ipcRenderer.invoke("setTheme", theme);
	}
});

ipcRenderer.on("pageLoaded", (event, page, html) => {
	const mainElement = document.getElementById("main");
	if (!mainElement) return;

	mainElement.innerHTML = html;
	currentPage = page;

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
