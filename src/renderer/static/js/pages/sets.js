// Imports
import { DataTable } from "../components/dataTable.js";

// Event Listeners
document.addEventListener("pageLoad", (event) => {
	let { page, params } = event.detail;
	let { tableRows } = params;

	if (page !== "sets") return;

	let dataTable = /** @type {DataTable} */ (document.getElementById("lego-sets"));

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
			header: "Release Year",
			accessorKey: "releaseYear"
		},
		{
			header: "Piece Count",
			accessorKey: "pieceCount"
		},
		{
			header: "Minifig Count",
			accessorKey: "minifigCount"
		},
		{
			header: "Amount",
			accessorKey: "legoSetCount",
			meta: {
				editable: true,
				type: "number",
				min: 1,
				onChange: (rowIndex, value) => {
					if (typeof value !== "number") return;

					window.electronAPI.changeSetCount(rowIndex, value);
				}
			},
		}
	];

	dataTable.data = tableRows;

	dataTable.renderTable();
});
