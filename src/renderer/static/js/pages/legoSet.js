// Imports
import { DataTable } from "../components/dataTable.js";
/** @import {LegoSetsTableRow, LegoSetTableRow} from "../../../../types.js" */
/** @import { CompoundLegoSetPiece } from "../../../../models.js" */


// Types
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
	/** @type {LegoSetsTableRow[]} */
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
				min: (row) => {
					return /** @type {number} */ (row.original.min);
				},
				onChange: async ({ value: amountFound, element, row }) => {
					if (
						typeof amountFound !== "number"
						|| Number.isNaN(amountFound)
						|| amountFound < 0
					) amountFound = 0;

					let input = element.getElementsByTagName("input")[0];

					let min = Number.parseInt(input.min);
					if (amountFound < min) amountFound = min;

					input.value = amountFound.toString();

					let databaseId = /** @type {number} */ (row.original.databaseId);
					let legoPieceChanges = await window.electronAPI.changeAmountFound(legoSetId, databaseId, amountFound);

					for (let legoPieceChange of legoPieceChanges) {
						let tr = /** @type {HTMLTableRowElement} */ (document.querySelector(`[data-row-id="${legoPieceChange.databaseId}"]`));
						let amountFoundTd = /** @type {HTMLInputElement} */ (tr.querySelector(".amountFound input"));
						let amountLeftTd = /** @type {HTMLTableCellElement} */ (tr.getElementsByClassName("amountLeft")[0]);
						let completionTd = /** @type {HTMLTableCellElement} */ (tr.getElementsByClassName("completion")[0]);

						amountFoundTd.value = legoPieceChange.amountFound.toString();

						let min = 0;
						let compoundLegoPieceChange = /** @type {CompoundLegoSetPiece} */ (legoPieceChange);
						if (Array.isArray(compoundLegoPieceChange.componentLegoSetPieces)) {
							let maxStickeredPiecesFound = 0;

							for (let componentLegoSetPiece of compoundLegoPieceChange.componentLegoSetPieces) {
								let totalStickeredPiecesFound = 0;
								for (let stickeredLegoSetPiece of componentLegoSetPiece.stickeredLegoSetPieces) {
									totalStickeredPiecesFound += stickeredLegoSetPiece.amountFound;
								}

								if (totalStickeredPiecesFound > maxStickeredPiecesFound) maxStickeredPiecesFound = totalStickeredPiecesFound;
							}

							min = maxStickeredPiecesFound;
						} else {
							let totalStickeredPiecesFound = 0;
							for (let stickeredLegoSetPiece of legoPieceChange.stickeredLegoSetPieces) {
								totalStickeredPiecesFound += stickeredLegoSetPiece.amountFound;
							}
							min = totalStickeredPiecesFound;
						}
						amountFoundTd.min = min.toString();

						let amountLeft = legoPieceChange.amountNeeded - legoPieceChange.amountFound;
						amountLeftTd.innerText = amountLeft.toString();

						completionTd.className = "completion";
						completionTd.classList.add(getCompletion(legoPieceChange.amountFound, legoPieceChange.amountNeeded));
					}
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
			id: "brickLinkName",
			header: "Name",
			accessorKey: "brickLinkName"
		},
		{
			id: "brickLinkId",
			header: "Id",
			accessorKey: "brickLinkId"
		},
		{
			id: "color",
			header: "Color",
			accessorFn: (row) => {
				const legoSetRow = /** @type {LegoSetTableRow} */ (row);
				return legoSetRow.color?.brickLinkName || "";
			}
		},
		{
			id: "brickLinkCategory",
			header: "Category",
			accessorKey: "brickLinkCategory"
		}
	];

	dataTable.data = tableRows;

	dataTable.renderTable();
});
