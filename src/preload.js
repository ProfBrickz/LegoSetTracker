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
	}
});

ipcRenderer.on("pageLoaded", (event, { page, html }) => {
	const mainElement = document.getElementById("main");
	if (!mainElement) return;

	mainElement.innerHTML = html;
	currentPage = page;

	// Update active button
	const navButtons = document.querySelectorAll(".nav-link");
	navButtons.forEach(button => {
		button.classList.remove("active");
	});

	// Add active class to current page button
	const currentButton = document.querySelector(`.nav-link[onclick*="${page}"]`);
	if (currentButton) {
		currentButton.classList.add("active");
	}

	// Notify that page has changed so icons can be recreated
	document.dispatchEvent(new CustomEvent("pageChanged"));
});
