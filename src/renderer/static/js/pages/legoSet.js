// Imports
import { DataTable } from "../components/dataTable.js";
/** @import { TableRow } from "../components/dataTable.js" */


/**
 * @typedef {"not-started" | "incomplete" | "complete" | "extra"} CompletionStatus
 */

// Functions
/**
 * @param {number} amountFound
 * @param {number} amountNeeded
 * @returns {CompletionStatus}
 */
function getCompletion(amountFound, amountNeeded) {
	/** @type {CompletionStatus} */
	let completion = "not-started";

	if (amountFound <= 0) {
		completion = "not-started";
	} else if (amountFound < amountNeeded) {
		completion = "incomplete";
	} else if (amountFound == amountNeeded) {
		completion = "complete";
	} else if (amountFound > amountNeeded) {
		completion = "extra";
	}

	return completion;
}


// Event Listeners
document.addEventListener("pageLoad", (event) => {
	/** @type {string} */
	let page = event.detail.page;
	/** @type {import("../../../../types.js").LegoSetsTableRow[]} */
	let tableRows = event.detail.params.tableRows;
	/** @type {number} */
	let legoSetId = event.detail.params.legoSetId;

	if (page !== "lego-set") return;

	let dataTable = /** @type {DataTable} */ (document.getElementById("lego-set"));

	dataTable.columns = [
		{
			id: "completion",
			header: "Completion Status",
			cell: () => {
				let fragment = document.createDocumentFragment();

				let notStarted = document.createElement("i");
				notStarted.classList.add("not-started");
				notStarted.dataset.lucide = "x";
				fragment.appendChild(notStarted);

				let incomplete = document.createElement("i");
				incomplete.classList.add("incomplete");
				incomplete.dataset.lucide = "x";
				fragment.appendChild(incomplete);

				let complete = document.createElement("i");
				complete.classList.add("complete");
				complete.dataset.lucide = "check";
				fragment.appendChild(complete);

				let extra = document.createElement("i");
				extra.classList.add("extra");
				extra.dataset.lucide = "plus";
				fragment.appendChild(extra);

				return fragment;
			},
			meta: {
				type: "function",
				classList: ({ row }) => {
					let amountFound = /** @type {number} */ (row.original.amountFound);
					let amountNeeded = /** @type {number} */ (row.original.amountNeeded);

					return [getCompletion(amountFound, amountNeeded)];
				}
			}
		},
		{
			id: "image",
			header: "Image",
			accessorKey: "image",
			meta: {
				type: "image"
			}
		},
		{
			id: "amountFound",
			header: "Amount Found",
			accessorKey: "amountFound",
			meta: {
				type: "number",
				min: 0,
				onChange: ({ value: amountFound, element, row }) => {
					if (typeof amountFound !== "number") return;

					let pieceId = /** @type {number} */ (row.original.pieceId);
					let amountNeeded = /** @type {number} */ (row.original.amountNeeded);
					let amountLeft = amountNeeded - amountFound;

					let tr = /** @type {HTMLTableRowElement} */ (element.parentElement);
					let amountLeftTd = /** @type {HTMLTableCellElement} */ (tr.getElementsByClassName("amountLeft")[0]);
					let completionTd = /** @type {HTMLTableCellElement} */ (tr.getElementsByClassName("completion")[0]);

					amountLeftTd.innerText = amountLeft.toString();

					completionTd.className = "completion";
					completionTd.classList.add(getCompletion(amountFound, amountNeeded));

					window.electronAPI.changeAmountFound(legoSetId, pieceId, amountFound);
				}
			}
		},
		{
			id: "amountLeft",
			header: "Amount Left",
			accessorKey: "amountLeft"
		},
		{
			id: "amountNeeded",
			header: "Amount Needed",
			accessorKey: "amountNeeded"
		},
		{
			id: "bricklinkName",
			header: "Name",
			accessorKey: "bricklinkName"
		},
		{
			id: "bricklinkId",
			header: "Id",
			accessorKey: "bricklinkId"
		},
		{
			id: "color",
			header: "Color",
			accessorFn: (row) => {
				const legoSetRow = /** @type {import("../../../../types.js").LegoSetTableRow} */ (row);
				return legoSetRow.color?.bricklinkName || "";
			}
		},
		{
			id: "bricklinkCategory",
			header: "Category",
			accessorKey: "bricklinkCategory"
		}
	];

	dataTable.data = tableRows;

	dataTable.renderTable();
});
