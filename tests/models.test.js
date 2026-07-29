// Imports
import { beforeEach, describe, expect, jest, test } from "@jest/globals";
import { LegoColors } from "../src/dataMaps.js";
import { CompoundLegoPiece, CompoundLegoSetPiece, LegoPiece, LegoSetPiece, StickeredLegoPiece, StickeredLegoSetPiece } from "../src/models.js";
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

describe("CompoundLegoSetPiece", () => {
	/** @type {LegoPiece} */
	let component1LegoPiece;
	/** @type {LegoPiece} */
	let component2LegoPiece;
	/** @type {CompoundLegoPiece} */
	let compoundLegoPiece;

	/** @type {LegoSetPiece} */
	let component1LegoSetPiece;
	/** @type {LegoSetPiece} */
	let component2LegoSetPiece;
	/** @type {CompoundLegoSetPiece} */
	let compoundLegoSetPiece;

	beforeEach(() => {
		component1LegoPiece = new LegoPiece(0, "", "", null, "");
		component2LegoPiece = new LegoPiece(0, "", "", null, "");
		compoundLegoPiece = new CompoundLegoPiece(0, "", "", null, "", [component1LegoPiece, component2LegoPiece]);

		component1LegoSetPiece = new LegoSetPiece(0, component1LegoPiece, 5, 0);
		component2LegoSetPiece = new LegoSetPiece(0, component2LegoPiece, 3, 0);
		compoundLegoSetPiece = new CompoundLegoSetPiece(0, compoundLegoPiece, 3, 0, [component1LegoSetPiece, component2LegoSetPiece]);
	});

	describe("syncAmountFound", () => {
		test("Add", () => {
			compoundLegoSetPiece.syncAmountFound(1);
			expect(compoundLegoSetPiece.amountFound).toEqual(1);
			expect(component1LegoSetPiece.amountFound).toEqual(1);
			expect(component2LegoSetPiece.amountFound).toEqual(1);

			compoundLegoSetPiece.syncAmountFound(2);
			expect(compoundLegoSetPiece.amountFound).toEqual(2);
			expect(component1LegoSetPiece.amountFound).toEqual(2);
			expect(component2LegoSetPiece.amountFound).toEqual(2);
		});

		test("Subtract", () => {
			compoundLegoSetPiece.amountFound = 1;
			component1LegoSetPiece.amountFound = 1;
			component2LegoSetPiece.amountFound = 1;

			compoundLegoSetPiece.syncAmountFound(2);
			expect(compoundLegoSetPiece.amountFound).toEqual(2);
			expect(component1LegoSetPiece.amountFound).toEqual(2);
			expect(component2LegoSetPiece.amountFound).toEqual(2);

			compoundLegoSetPiece.syncAmountFound(1);
			expect(compoundLegoSetPiece.amountFound).toEqual(1);
			expect(component1LegoSetPiece.amountFound).toEqual(1);
			expect(component2LegoSetPiece.amountFound).toEqual(1);
		});

		test("Add, Subtract", () => {
			compoundLegoSetPiece.syncAmountFound(1);
			expect(compoundLegoSetPiece.amountFound).toEqual(1);
			expect(component1LegoSetPiece.amountFound).toEqual(1);
			expect(component2LegoSetPiece.amountFound).toEqual(1);

			compoundLegoSetPiece.syncAmountFound(2);
			expect(compoundLegoSetPiece.amountFound).toEqual(2);
			expect(component1LegoSetPiece.amountFound).toEqual(2);
			expect(component2LegoSetPiece.amountFound).toEqual(2);

			compoundLegoSetPiece.syncAmountFound(1);
			expect(compoundLegoSetPiece.amountFound).toEqual(1);
			expect(component1LegoSetPiece.amountFound).toEqual(1);
			expect(component2LegoSetPiece.amountFound).toEqual(1);
		});

		test("Add different", () => {
			compoundLegoSetPiece.amountFound = 0;
			component1LegoSetPiece.amountFound = 2;
			component2LegoSetPiece.amountFound = 2;

			compoundLegoSetPiece.syncAmountFound(1);
			expect(compoundLegoSetPiece.amountFound).toEqual(1);
			expect(component1LegoSetPiece.amountFound).toEqual(3);
			expect(component2LegoSetPiece.amountFound).toEqual(3);
		});
	});

	test("updateAmountFound", () => {
		compoundLegoSetPiece.updateAmountFound();
		expect(component1LegoSetPiece.amountFound).toEqual(0);
		expect(component2LegoSetPiece.amountFound).toEqual(0);
		expect(compoundLegoSetPiece.amountFound).toEqual(0);

		component1LegoSetPiece.amountFound = 1;
		compoundLegoSetPiece.updateAmountFound();
		expect(component1LegoSetPiece.amountFound).toEqual(1);
		expect(component2LegoSetPiece.amountFound).toEqual(0);
		expect(compoundLegoSetPiece.amountFound).toEqual(0);

		component2LegoSetPiece.amountFound = 1;
		compoundLegoSetPiece.updateAmountFound();
		expect(component1LegoSetPiece.amountFound).toEqual(1);
		expect(component2LegoSetPiece.amountFound).toEqual(1);
		expect(compoundLegoSetPiece.amountFound).toEqual(1);

		component1LegoSetPiece.amountFound = 5;
		component2LegoSetPiece.amountFound = 5;
		compoundLegoSetPiece.updateAmountFound();
		expect(component1LegoSetPiece.amountFound).toEqual(5);
		expect(component2LegoSetPiece.amountFound).toEqual(5);
		expect(compoundLegoSetPiece.amountFound).toEqual(5);

		component1LegoSetPiece.amountFound = 0;
		compoundLegoSetPiece.updateAmountFound();
		expect(component1LegoSetPiece.amountFound).toEqual(0);
		expect(component2LegoSetPiece.amountFound).toEqual(5);
		expect(compoundLegoSetPiece.amountFound).toEqual(0);

		component1LegoSetPiece.amountFound = 5;
		component2LegoSetPiece.amountFound = 5;
		compoundLegoSetPiece.legoSetComponents = [];
		compoundLegoSetPiece.updateAmountFound();
		expect(compoundLegoSetPiece.amountFound).toEqual(0);
	});
});
