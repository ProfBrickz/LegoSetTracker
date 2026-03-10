// Imports
import { DataTable } from "./components/dataTable.js";

// Event Listeners
document.addEventListener("pageLoad", (event) => {
	let { page } = event.detail;

	if (page !== "add-set") return;

	let setSearchForm = /** @type {HTMLFormElement} */ (document.getElementById("set-search-form"));
	setSearchForm.onsubmit = (event) => {
		event.preventDefault();
		window.electronAPI.searchLegoSets((searchResults) => {
			dataTable.data = searchResults;
			dataTable.renderTable();
		});
	};

	let dataTable = /** @type {DataTable} */ (document.getElementById("lego-set-search-results"));

	dataTable.columns = [
		{
			header: "Image",
			accessorKey: "image",
			meta: {
				type: "image"
			}
		},
		{
			header: "Name",
			accessorKey: "name"
		},
		{
			header: "Set Number",
			accessorKey: "setNumber"
		},
		{
			header: "Theme",
			accessorKey: "theme"
		},
		{
			header: "Add",
			cell: ({ row }) => {
				let addButton = document.createElement("button");
				addButton.innerText = "Add Set";

				addButton.onclick = () => {
					window.electronAPI.addSet(row.getValue("setNumber"));
				};

				return addButton;
			},
			meta: {
				type: "function"
			}
		}
	];

	// window.electronAPI.onSearchResults(searchResults, (searchResults) => {
	// dataTable.data = searchResults;
	// dataTable.renderTable();
	// });
});
