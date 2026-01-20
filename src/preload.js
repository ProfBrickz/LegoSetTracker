// Imports
const { contextBridge, ipcRenderer } = require('electron');


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

ipcRenderer.on("pageLoaded", (event, { html }) => {
	const mainElement = document.getElementById("main");
	if (!mainElement) return;

	mainElement.innerHTML = html;
});
