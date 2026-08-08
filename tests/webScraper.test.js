// Imports
import { afterEach, beforeEach, describe, expect, jest, test } from "@jest/globals";
import fs from "fs";
import { LegoColors } from "../src/dataMaps.js";
import { LegoColor, LegoPiece } from "../src/models.js";
import WebScraper from "../src/webScraper.js";
import { getRelativeFilePath, readTextFile } from "./testFunctions.js";
/** @import { LegoSetInfo, LegoSetPieceInfo, LegoSetPiecesInfo, LegoSetSearchResult, WebLegoColor, WebLegoSetTheme } from "../src/types.js" */


// Variables
/** @type {LegoColors} */
let colors;
/** @type {LegoPiece[]} */
let legoPieces;
/** @type {WebScraper} */
let webScraper;

/** @type {jest.SpiedFunction<fetch>} */
let fetchMock;


// Setup / Teardown
beforeEach(() => {
	fetchMock = jest.spyOn(global, "fetch");
	colors = new LegoColors();
	legoPieces = [];
	webScraper = new WebScraper(colors);
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

		let document = await webScraper.getWebpage(url);

		expect(fetch).toHaveBeenCalledWith(url, expect.objectContaining({ headers: expect.any(Headers) }));
		// Verify that the document is a valid HTML document and contains
		expect(document.constructor.name).toEqual("Document");
		expect(document.documentElement.tagName).toEqual("HTML");
		expect(document.querySelector("h1")?.textContent).toBe("Hello World!");
	});

	test("Throws error when the response is not ok", async () => {
		// Mock fetch to return a failed response
		fetchMock.mockResolvedValue(new Response("", { status: 404, statusText: "Not Found" }));

		// Expect an error to be thrown when the response is
		await expect(webScraper.getWebpage("https://example.com")).rejects.toThrow(
			new Error("Failed to fetch https://example.com: 404 Not Found")
		);
	});
});

describe("getColors", () => {
	test("Returns an array of colors", async () => {
		/** @type {WebLegoColor[]} */
		let result = [
			{ brickLinkId: 1, brickLinkName: "White", legoId: 1, legoName: "White" },
			{ brickLinkId: 5, brickLinkName: "Red", legoId: 21, legoName: "Bright Red" },
			{ brickLinkId: 48, brickLinkName: "Sand Green", legoId: 151, legoName: "Sand Green" },
			{ brickLinkId: 12, brickLinkName: "Trans-Clear", legoId: 40, legoName: "Transparent" },
			{ brickLinkId: 17, brickLinkName: "Trans-Red", legoId: 41, legoName: "Tr. Red" },
			{ brickLinkId: 14, brickLinkName: "Trans-Dark Blue", legoId: 43, legoName: "Tr. Blue" },
			{ brickLinkId: 122, brickLinkName: "Chrome Black", legoId: null, legoName: null },
		];

		let html = readTextFile("./fixtures/getColors/colors.html");
		// Mock fetch to return the HTML content of the fixture
		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let colors = await webScraper.getColors();

		// Verify that the returned value is the correct array of colors
		expect(Array.isArray(colors)).toEqual(true);
		expect(colors).toHaveLength(7);
		expect(colors).toEqual(result);
	});

	test("Throws an error", async () => {
		// Mock the fetch function to simulate a network error
		fetchMock.mockResolvedValue(new Response("", { status: 404, statusText: "Not Found" }));

		// Expect an error to be thrown when calling getColors
		await expect(webScraper.getColors()).rejects.toThrow(
			"Failed to fetch https://v2.bricklink.com/en-us/catalog/color-guide: 404 Not Found"
		);
	});

	test("Returns an empty array", async () => {
		let html = readTextFile("./fixtures/getColors/empty.html");

		// Mock the fetch function to return a response with the
		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let colors = await webScraper.getColors();

		// Verify that the function returns an empty array
		expect(colors).toBeInstanceOf(Array);
		expect(colors).toHaveLength(0);
		expect(colors).toEqual([]);
	});
});

describe("getLegoSetThemes", () => {
	test("Returns an Map of categories with id's and name's", async () => {
		/** @type {WebLegoSetTheme[]} */
		let result = [
			{ brickLinkId: "143", brickLinkName: "(Other)" },
			{ brickLinkId: "516", brickLinkName: "4 Juniors" },
			{ brickLinkId: "178", brickLinkName: "Jack Stone" },
			{ brickLinkId: "61", brickLinkName: "Pirates" },
			{ brickLinkId: "469", brickLinkName: "Spider-Man" },
			{ brickLinkId: "609", brickLinkName: "Agents" },
			{ brickLinkId: "1370", brickLinkName: "Education" },
			{ brickLinkId: "166", brickLinkName: "Educational & Dacta" },
			{ brickLinkId: "167", brickLinkName: "DUPLO" },
			{ brickLinkId: "173", brickLinkName: "Action Wheelers" },
			{ brickLinkId: "612", brickLinkName: "Town" },
			{ brickLinkId: "325", brickLinkName: "Airport" },
		];

		let html = readTextFile("./fixtures/getLegoSetThemes/success.html");

		// Mock fetch to return the HTML content of the fixture
		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetThemes = await webScraper.getLegoSetThemes();
		// Verify that the returned value is the correct array of colors
		expect(Array.isArray(legoSetThemes)).toEqual(true);
		expect(legoSetThemes).toHaveLength(12);
		expect(legoSetThemes).toEqual(result);
	});
});

describe("searchLegoSets", () => {
	test("Returns an array of Lego set search results", async () => {
		/** @type {LegoSetSearchResult[]} */
		let result = [
			{
				name: "Star Destroyer",
				setNumber: "75033-1",
				themeId: "258",
			}, {
				name: "Imperial Star Destroyer",
				setNumber: "75055-1",
				themeId: "258",
			}, {
				name: "First Order Star Destroyer",
				setNumber: "75190-1",
				themeId: "923",
			}, {
				name: "First Order Star Destroyer - Mini polybag",
				setNumber: "30277-1",
				themeId: "858",
			}, {
				name: "Star Destroyer + TIE Fighter - Mini foil pack",
				setNumber: "911510-1",
				themeId: "258",
			}, {
				name: "Mini Star Destroyer - Star Wars Celebration Anaheim 2015",
				setNumber: "CELEB2015SD-1",
				themeId: "983",
			}, {
				name: "Advent Calendar 2015, Star Wars (Day 11) - Star Destroyer",
				setNumber: "75097-12",
				themeId: "65",
			},
		];

		let json = readTextFile("./fixtures/searchLegoSets/success.json");
		fetchMock.mockResolvedValue(new Response(json, { status: 200 }));

		let searchResults = await webScraper.searchLegoSets(
			"destroyer",
			{ themeId: "65", startYear: 2014, endYear: 2017 }
		);

		expect(searchResults).toHaveLength(7);
		expect(searchResults).toEqual(result);
	});

	test("Throws error when not ok", async () => {
		fetchMock.mockResolvedValue(new Response("", { status: 400, statusText: "Bad Request" }));

		await expect(webScraper.searchLegoSets("")).rejects.toThrow("Failed to fetch data");
	});

	test("Invalid parameters", async () => {
		let json = readTextFile("./fixtures/searchLegoSets/error.json");
		fetchMock.mockResolvedValue(new Response(json, { status: 200 }));

		// let searchResults = await webScraper.searchLegoSets("");
		await expect(webScraper.searchLegoSets("")).rejects.toThrow("Query keyword is not specified!");
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
			themeId: "841",
			releaseYear: 2015,
			pieceCount: 46,
			minifigCount: 2
		};

		let html = readTextFile("./fixtures/getLegoSetInfo/success.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetInfo = await webScraper.getLegoSetInfo(result.setNumber);

		expect(legoSetInfo).toEqual(result);
	});

	test("Not Found", async () => {
		let setNumber = "10679-1";

		let html = readTextFile("./fixtures/getLegoSetInfo/not-found.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		await expect(webScraper.getLegoSetInfo(setNumber)).rejects.toThrow(`Could not find the set ${setNumber}`);
	});
});

describe("getLegoSetPieces", () => {
	beforeEach(() => {
		colors.set(0, new LegoColor(0, 1, "White", 1, "White"));
		colors.set(1, new LegoColor(1, 85, "Dark Blueish Gray", 199, "Dark Stone Grey"),);
		colors.set(2, new LegoColor(2, 11, "Black", 26, "Black"),);
		colors.set(3, new LegoColor(3, 2, "Tan", 5, "Brick Yellow"),);
		colors.set(4, new LegoColor(4, 14, "Trans-Dark Blue", 43, "Tr. Blue"));
	});

	test("Successfully fetch Lego set pieces", async () => {
		/** @type {LegoSetPiecesInfo} */
		let result = {
			normalPieces: [
				{
					brickLinkId: "4738a",
					brickLinkName: "Container, Treasure Chest Bottom with Slots in Back",
					color: colors.get(2) || null,
					brickLinkCategory: "Container",
					amountNeeded: 1,
					amountFound: 0
				},
				{
					brickLinkId: "4739a",
					brickLinkName: "Container, Treasure Chest Lid Curved with Thick Hinge",
					color: colors.get(2) || null,
					brickLinkCategory: "Container",
					amountNeeded: 1,
					amountFound: 0
				},
				{
					brickLinkId: "92338",
					brickLinkName: "Dark Bluish Gray Chain 5 Links",
					color: colors.get(1) || null,
					brickLinkCategory: "Chain",
					amountNeeded: 1,
					amountFound: 0
				},
				{
					brickLinkId: "3068pb0906",
					brickLinkName: "Tile 2 x 2 with Map Blue Water, Lime Land, Sailing Ship, Treasure Chest and Red 'X' Pattern",
					color: colors.get(3) || null,
					brickLinkCategory: "Tile, Decorated",
					amountNeeded: 1,
					amountFound: 0
				},
				{
					brickLinkId: "30153",
					brickLinkName: "Rock 1 x 1 Jewel 24 Facet",
					color: colors.get(4) || null,
					brickLinkCategory: "Rock",
					amountNeeded: 2,
					amountFound: 0
				}
			],
			minifigs: [
				{
					brickLinkId: "pi146",
					brickLinkName: "Pirate Blue Jacket, Black Leg with Peg Leg, Black Pirate Hat with Skull",
					color: null,
					brickLinkCategory: "Pirates",
					amountNeeded: 1,
					amountFound: 0
				},
				{
					brickLinkId: "gen067",
					brickLinkName: "Skeleton - Standard Skull, Floppy Arms, Red Bandana with Double Tail in Back",
					color: null,
					brickLinkCategory: "Pirates",
					amountNeeded: 10,
					amountFound: 0
				}
			],
			extraPieces: [
				{
					brickLinkId: "92338",
					brickLinkName: "Dark Bluish Gray Chain 5 Links",
					color: colors.get(1) || null,
					brickLinkCategory: "Chain",
					amountNeeded: 1,
					amountFound: 0
				}
			],
			counterparts: [
				{
					brickLinkId: "4738ac01",
					brickLinkName: "Container, Treasure Chest with Slots in Back and (Same Color) Thick Hinge Curved Lid (4738a / 4739a)",
					color: colors.get(2) || null,
					brickLinkCategory: "Container",
					amountNeeded: 1,
					amountFound: 0
				}
			]
		};

		let html = readTextFile("./fixtures/getLegoSetPieces/success.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetPieceInfo = await webScraper.getLegoSetPieces("10679-1");

		expect(legoSetPieceInfo.normalPieces[0].brickLinkId).toEqual(result.normalPieces[0].brickLinkId);
		expect(legoSetPieceInfo.normalPieces[0].brickLinkName).toEqual(result.normalPieces[0].brickLinkName);
		expect(legoSetPieceInfo.normalPieces[0].color).toBe(result.normalPieces[0].color);
		expect(legoSetPieceInfo.normalPieces[0].brickLinkCategory).toEqual(result.normalPieces[0].brickLinkCategory);
		expect(legoSetPieceInfo).toEqual(result);
	});

	test("Return empty array if no pieces are found", async () => {
		/** @type {LegoSetPiecesInfo} */
		let result = {
			normalPieces: [],
			minifigs: [],
			extraPieces: [],
			counterparts: []
		};

		let html = readTextFile("./fixtures/getLegoSetPieces/empty.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetPieceInfo = await webScraper.getLegoSetPieces("10679-1");

		expect(legoSetPieceInfo).toEqual(result);
	});
});

describe("downloadImage", () => {
	let sourceFile = getRelativeFilePath("./fixtures/downloadImage/image.jpg");
	let downloadFile = getRelativeFilePath("./image.jpg");
	let imageData = fs.readFileSync(sourceFile);
	let url = "https://example.com/image.jpg";

	afterEach(() => {
		fs.rmSync(downloadFile, { force: true });
	});

	test("Downloads image", async () => {
		fetchMock.mockResolvedValue(new Response(imageData, { status: 200 }));

		// Verify that the file does not exist before downloading it
		expect(fs.existsSync(downloadFile)).toEqual(false);

		await webScraper.downloadImage(url, downloadFile);

		// Verify that the file was downloaded
		expect(fetch).toHaveBeenCalledWith(url);
		expect(fs.existsSync(downloadFile)).toEqual(true);
		expect(fs.readFileSync(downloadFile)).toEqual(fs.readFileSync(sourceFile));
	});

	test("Does not download image if already exists", async () => {
		fs.writeFileSync(downloadFile, imageData);

		fetchMock.mockResolvedValue(new Response(imageData, { status: 200 }));

		// Check if file exists before downloading
		expect(fs.existsSync(downloadFile)).toEqual(true);
		await webScraper.downloadImage(url, downloadFile);

		// Check if file still exists after attempting to download again
		expect(fetch).not.toHaveBeenCalledWith(url);
		expect(fs.existsSync(downloadFile)).toEqual(true);
	});

	test("Throws error when the response is not ok", async () => {
		fetchMock.mockResolvedValue(new Response(imageData, { status: 404, statusText: "Not Found" }));

		await expect(webScraper.downloadImage("https://example.com/image.jpg", downloadFile)).rejects.toThrow(
			new Error("Failed to fetch https://example.com/image.jpg: 404 Not Found")
		);
	});

	test("Throws error when write fails", async () => {
		fetchMock.mockResolvedValue(new Response(imageData, { status: 200 }));

		// Mock fs.promises.writeFile to simulate a write failure
		let writeFileMock = jest.spyOn(fs.promises, "writeFile");
		writeFileMock.mockImplementation(() => {
			throw new Error("Write failed");
		});

		await expect(webScraper.downloadImage(url, downloadFile)).rejects.toThrow("Write failed");
	});
});
