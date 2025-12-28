// import { contextBridge, ipcRenderer } from 'electron';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld("electronAPI", {
	/**
	 * @param {string} page The page to navigate to.
	 * @param {Object} params The parameters for the page.
	 * @returns {void}
	 */
	navigate: (page, params = {}) => {
		ipcRenderer.send("navigate", { page, params });
	}
});
