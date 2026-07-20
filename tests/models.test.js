// Imports
import { beforeEach, describe, expect, jest, test } from "@jest/globals";
import { LegoColors } from "../src/dataMaps.js";
import { LegoPiece, LegoSetPiece, StickeredLegoPiece, StickeredLegoSetPiece } from "../src/models.js";
import WebScrapper from "../src/webScrapper.js";
/** @import { LegoSetInfo, LegoSetPieceInfo, LegoSetPiecesInfo, LegoSetSearchResult, WebLegoColor, WebLegoSetTheme } from "../src/types.js" */


// Variables
/** @type {LegoColors} */
let colors;
/** @type {LegoPiece[]} */
let legoPieces;
/** @type {WebScrapper} */
let webScrapper;

/** @type {jest.SpiedFunction<fetch>} */
let fetchMock;





// Tests
describe("StickeredLegoSetPiece", () => {
	/** @type {LegoPiece} */
	let baseLegoPiece;
	/** @type {StickeredLegoPiece} */
	let stickeredLegoPiece;
	/** @type {LegoSetPiece} */
	let baseLegoSetPiece;
	/** @type {StickeredLegoSetPiece} */
	let stickeredLegoSetPiece;

	beforeEach(() => {
		baseLegoPiece = new LegoPiece(0, "", "", null, "");
		stickeredLegoPiece = new StickeredLegoPiece(0, "", "", null, "", baseLegoPiece);
		baseLegoSetPiece = new LegoSetPiece(0, baseLegoPiece, 5, 0);
		stickeredLegoSetPiece = new StickeredLegoSetPiece(0, stickeredLegoPiece, baseLegoSetPiece, 3, 0);
	});

	describe("syncAmountFound", () => {
		test("Add", () => {
			stickeredLegoSetPiece.syncAmountFound(1);
			expect(stickeredLegoSetPiece.amountFound).toEqual(1);
			expect(baseLegoSetPiece.amountFound).toEqual(1);

			stickeredLegoSetPiece.syncAmountFound(2);
			expect(stickeredLegoSetPiece.amountFound).toEqual(2);
			expect(baseLegoSetPiece.amountFound).toEqual(2);
		});

		test("Subtract", () => {
			stickeredLegoSetPiece.amountFound = 1;
			baseLegoSetPiece.amountFound = 1;

			stickeredLegoSetPiece.syncAmountFound(2);
			expect(stickeredLegoSetPiece.amountFound).toEqual(2);
			expect(baseLegoSetPiece.amountFound).toEqual(2);

			stickeredLegoSetPiece.syncAmountFound(1);
			expect(stickeredLegoSetPiece.amountFound).toEqual(1);
			expect(baseLegoSetPiece.amountFound).toEqual(1);
		});

		test("Add, Subtract", () => {
			stickeredLegoSetPiece.syncAmountFound(1);
			expect(stickeredLegoSetPiece.amountFound).toEqual(1);
			expect(baseLegoSetPiece.amountFound).toEqual(1);

			stickeredLegoSetPiece.syncAmountFound(2);
			expect(stickeredLegoSetPiece.amountFound).toEqual(2);
			expect(baseLegoSetPiece.amountFound).toEqual(2);

			stickeredLegoSetPiece.syncAmountFound(1);
			expect(stickeredLegoSetPiece.amountFound).toEqual(1);
			expect(baseLegoSetPiece.amountFound).toEqual(1);
		});

		test("Add different", () => {
			stickeredLegoSetPiece.amountFound = 0;
			baseLegoSetPiece.amountFound = 2;

			stickeredLegoSetPiece.syncAmountFound(1);
			expect(stickeredLegoSetPiece.amountFound).toEqual(1);
			expect(baseLegoSetPiece.amountFound).toEqual(3);
		});
	});
});
