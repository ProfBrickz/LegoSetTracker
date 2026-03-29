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
				onChange: ({ row, value: setCount }) => {
					if (typeof setCount !== "number") return;
					let databaseId = /** @type {number} */ (row.original.databaseId);

					window.electronAPI.changeLegoSetCount(databaseId, setCount);
				}
			},
		},
		{
			header: "Buttons",
			cell: ({ row }) => {
				let fragment = document.createDocumentFragment();

				let viewButton = document.createElement("button");
				viewButton.innerText = "View";
				viewButton.onclick = () => window.electronAPI.loadPage(
					"lego-set",
					{ params: { databaseId: row.original.databaseId } }
				);
				fragment.appendChild(viewButton);

				let deleteButton = document.createElement("button");
				deleteButton.innerText = "Delete";
				deleteButton.classList.add("danger");
				deleteButton.onclick = async () => {
					let confirmDelete = confirm(`Are you sure you want to delete ${row.original.name}`);
					if (confirmDelete) {
						await window.electronAPI.deleteLegoSet(/** @type {number} */(row.original.databaseId));
						dataTable.deleteRow(row.index);
					}
				};
				fragment.appendChild(deleteButton);

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
