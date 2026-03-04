// Imports
import { DataTable } from "./components/dataTable.js";


document.addEventListener("pageLoad", (event) => {
	let { page } = event.detail;

	if (page !== "add-set") return;

	let dataTable = /** @type {DataTable} */ (document.getElementById("lego-set-search-results"));

	dataTable.columns = [
		{
			header: "Name",
			accessorKey: "name"
		},
		{
			header: "Set Number",
			accessorKey: "setNumber"
		},
		{
			header: "Theme ID",
			accessorKey: "themeId"
		},
		{
			header: "Theme",
			accessorKey: "theme"
		}
	];

	window.electronAPI.onSearchResults((event, searchResults) => {
		dataTable.data = searchResults;
		dataTable.renderTable();
	});
});
