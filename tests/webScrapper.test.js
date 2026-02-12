// Imports
import { afterEach, beforeEach, describe, expect, jest, test } from "@jest/globals";
import { LegoColor, LegoPiece, LegoSetPiece } from "../src/models.js";
import WebScrapper from "../src/webScrapper.js";
import { readTextFile } from "./testFunctions.js";
/** @import { LegoSetInfo, LegoSetPieceInfo } from "../src/types.js" */


// Variables
/** @type {LegoColor[]} */
let colors;
/** @type {LegoPiece[]} */
let legoPieces;
/** @type {WebScrapper} */
let webScrapper;

/** @type {jest.SpiedFunction<fetch>} */
let fetchMock;


// Setup / Teardown
beforeEach(() => {
	fetchMock = jest.spyOn(global, "fetch");
	colors = [];
	legoPieces = [];
	webScrapper = new WebScrapper(colors, legoPieces);
});

afterEach(() => {
	jest.resetAllMocks();
});


// Tests
describe("getWebpage", () => {
	test("Fetches the URL and returns a document on success", async () => {
		let html = readTextFile("./fixtures/getWebpage/success.html");
		let url = "https://example.com";

		// Mock fetch to return a successful response with HTML content
		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let document = await webScrapper.getWebpage(url);

		expect(fetch).toHaveBeenCalledWith(url);
		// Verify that the document is a valid HTML document and contains
		expect(document.constructor.name).toEqual("Document");
		expect(document.documentElement.tagName).toEqual("HTML");
		expect(document.querySelector("h1")?.textContent).toBe("Hello World!");
	});

	test("Throws error when the response is not ok", async () => {
		// Mock fetch to return a failed response
		fetchMock.mockResolvedValue(new Response("", { status: 404, statusText: "Not Found" }));

		// Expect an error to be thrown when the response is
		await expect(webScrapper.getWebpage("https://example.com")).rejects.toThrow(
			new Error("Failed to fetch https://example.com: 404 Not Found")
		);
	});
});

describe("getColors", () => {
	test("Returns an array of colors", async () => {
		let result = [
			new LegoColor(null, 1, "White", 1, "White"),
			new LegoColor(null, 5, "Red", 21, "Bright Red"),
			new LegoColor(null, 48, "Sand Green", 151, "Sand Green"),
			new LegoColor(null, 12, "Trans-Clear", 40, "Transparent"),
			new LegoColor(null, 17, "Trans-Red", 41, "Tr. Red"),
			new LegoColor(null, 14, "Trans-Dark Blue", 43, "Tr. Blue"),
			new LegoColor(null, 122, "Chrome Black", null, "")
		];

		let html = readTextFile("./fixtures/getColors/colors.html");

		// Mock fetch to return the HTML content of the fixture
		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let colors = await webScrapper.getColors();

		// Verify that the returned value is the correct array of colors
		expect(colors).toBeInstanceOf(Array);
		expect(colors.length).toEqual(7);
		expect(colors[0]).toBeInstanceOf(LegoColor);
		expect(colors).toEqual(result);
	});

	test("Throws an error", async () => {
		// Mock the fetch function to simulate a network error
		fetchMock.mockResolvedValue(new Response("", { status: 404, statusText: "Not Found" }));

		// Expect an error to be thrown when calling getColors
		await expect(webScrapper.getColors()).rejects.toThrow(
			"Failed to fetch https://v2.bricklink.com/en-us/catalog/color-guide: 404 Not Found"
		);
	});

	test("Returns an empty array", async () => {
		let html = readTextFile("./fixtures/getColors/empty.html");

		// Mock the fetch function to return a response with the
		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let colors = await webScrapper.getColors();

		// Verify that the function returns an empty array
		expect(colors).toBeInstanceOf(Array);
		expect(colors.length).toEqual(0);
		expect(colors).toEqual([]);
	});
});

describe("getMinifigPieces", () => {
	// TODO: Implement after MVP
	test.todo("Implement after MVP");
});

describe("getCompositePiece", () => {
	// TODO: Implement after MVP
	test.todo("Implement after MVP");
});

describe("getLegoSetInfo", () => {
	test("Successfully retrieves LEGO set info", async () => {
		/** @type {LegoSetInfo} */
		let result = {
			setNumber: "10679-1",
			name: "Pirate Treasure Hunt",
			theme: "Juniors, Pirates, Pirates III",
			releaseYear: 2015,
			pieceCount: 46,
			minifigCount: 2
		};

		let html = readTextFile("./fixtures/getLegoSetInfo/success.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetInfo = await webScrapper.getLegoSetInfo(result.setNumber);

		expect(legoSetInfo).toEqual(result);
	});
});

describe("getLegoSetPieces", () => {
	beforeEach(() => {
		colors.push(
			new LegoColor(1, 1, "White", 1, "White"),
			new LegoColor(2, 85, "Dark Blueish Gray", 199, "Dark Stone Grey"),
			new LegoColor(3, 11, "Black", 26, "Black"),
			new LegoColor(4, 2, "Tan", 5, "Brick Yellow"),
			new LegoColor(5, 14, "Trans-Dark Blue", 43, "Tr. Blue")
		);
	});

	test("Successfully fetch Lego set pieces", async () => {
		/** @type {LegoSetPieceInfo} */
		let result = {
			normalPieces: [
				new LegoSetPiece(
					null,
					new LegoPiece(null, "4738a", "Container, Treasure Chest Bottom with Slots in Back", colors[2], "Container"),
					1
				),
				new LegoSetPiece(
					null,
					new LegoPiece(null, "4739a", "Container, Treasure Chest Lid Curved with Thick Hinge", colors[2], "Container"),
					1
				),
				new LegoSetPiece(
					null,
					new LegoPiece(null, "92338", "Chain 5 Links", colors[1], "Chain"),
					1
				),
				new LegoSetPiece(
					null,
					new LegoPiece(null, "3068pb0906", "Tile 2 x 2 with Map Blue Water, Lime Land, Sailing Ship, Treasure Chest and Red 'X' Pattern", colors[3], "Tile, Decorated"),
					1
				),
				new LegoSetPiece(
					null,
					new LegoPiece(null, "30153", "Rock 1 x 1 Jewel 24 Facet", colors[4], "Rock"),
					2
				),
			],
			minifigs: [
				new LegoSetPiece(
					null,
					new LegoPiece(null, "pi146", "Pirate Blue Jacket, Black Leg with Peg Leg, Black Pirate Hat with Skull", null, "Pirates"),
					1
				),
				new LegoSetPiece(
					null,
					new LegoPiece(null, "gen067", "Skeleton - Standard Skull, Floppy Arms, Red Bandana with Double Tail in Back", null, "Pirates"),
					10
				)
			],
			extraPieces: [
				new LegoSetPiece(
					null,
					new LegoPiece(null, "92338", "Chain 5 Links", colors[1], "Chain"),
					1
				)
			],
			counterparts: [
				new LegoSetPiece(
					null,
					new LegoPiece(null, "4738ac01", "Container, Treasure Chest with Slots in Back and (Same Color) Thick Hinge Curved Lid (4738a / 4739a)", colors[2], "Container"),
					1
				)
			]
		};

		let html = readTextFile("./fixtures/getLegoSetPieces/success.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetPieceInfo = await webScrapper.getLegoSetPieces("10679-1");

		expect(legoSetPieceInfo.normalPieces[0].bricklinkId).toEqual(result.normalPieces[0].bricklinkId);
		expect(legoSetPieceInfo.normalPieces[0].bricklinkName).toEqual(result.normalPieces[0].bricklinkName);
		expect(legoSetPieceInfo.normalPieces[0].color).toBe(result.normalPieces[0].color);
		expect(legoSetPieceInfo.normalPieces[0].bricklinkCategory).toEqual(result.normalPieces[0].bricklinkCategory);
		expect(legoSetPieceInfo).toEqual(result);
	});
});

describe("getLegoSet", () => {
	// TODO: Implement
	test.todo("Implement");
});
