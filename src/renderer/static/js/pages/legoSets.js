// Imports
import { DataTable } from "../components/dataTable.js";


// Event Listeners
document.addEventListener("pageLoad", (event) => {
	/** @type {string} */
	let page = event.detail.page;
	/** @type {import("../../../../types.js").LegoSetsTableRow[]} */
	let tableRows = event.detail.params.tableRows;

	if (page !== "lego-sets") return;

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
				type: "number",
				min: 1,
				onChange: ({ rowIndex, value: setCount }) => {
					if (typeof setCount !== "number") return;

					window.electronAPI.changeSetCount(rowIndex, setCount);
				}
			},
		},
		{
			header: "Buttons",
			cell: ({ row }) => {
				let fragment = document.createDocumentFragment();

				let viewButton = document.createElement("button");
				viewButton.innerText = "View";

				viewButton.onclick = () => window.electronAPI.loadPage("lego-set", { params: { databaseId: row.original.databaseId } });
				fragment.appendChild(viewButton);

				return fragment;
			},
			meta: {
				type: "function"
			}
		}
	];

	dataTable.data = tableRows;

	dataTable.renderTable();
});
