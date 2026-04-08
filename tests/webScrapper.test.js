// Imports
import { afterEach, beforeEach, describe, expect, jest, test } from "@jest/globals";
import fs from "fs";
import { LegoColor, LegoPiece, LegoSetPiece } from "../src/models.js";
import WebScrapper from "../src/webScrapper.js";
import { getRelativeFilePath, readTextFile } from "./testFunctions.js";
/** @import { LegoSetInfo, LegoSetPieceInfo, LegoSetSearchResult } from "../src/types.js" */


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

describe("getLegoSetThemes", () => {
	test("Returns an Map of categories with id's and name's", async () => {
		/** @type {Map<string, string>} */
		let result = new Map([
			["143", "(Other)"],
			["516", "4 Juniors"],
			["516.178", "Jack Stone"],
			["516.61", "Pirates"],
			["516.469", "Spider-Man"],
			["609", "Agents"],
			["1370", "Education"],
			["166", "Educational & Dacta"],
			["166.167", "DUPLO"],
			["166.167.173", "Action Wheelers"],
			["166.167.612", "Town"],
			["166.167.612.325", "Airport"]
		]);

		let html = readTextFile("./fixtures/getLegoSetThemes/success.html");

		// Mock fetch to return the HTML content of the fixture
		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetCategories = await webScrapper.getLegoSetThemes();

		// Verify that the returned value is the correct array of colors
		expect(legoSetCategories).toBeInstanceOf(Map);
		// expect(legoSetCategories.length).toEqual(7);
		// expect(legoSetCategories[0]).toBeInstanceOf(LegoColor);
		expect(legoSetCategories).toEqual(result);
	});
});

describe("searchLegoSets", () => {
	test("Returns an array of Lego set search results", async () => {
		/** @type {LegoSetSearchResult[]} */
		let result = [
			{
				name: "Star Destroyer",
				setNumber: "75033-1",
				themeId: "65.806.258",
			}, {
				name: "Imperial Star Destroyer",
				setNumber: "75055-1",
				themeId: "65.258",
			}, {
				name: "First Order Star Destroyer",
				setNumber: "75190-1",
				themeId: "65.923",
			}, {
				name: "First Order Star Destroyer - Mini polybag",
				setNumber: "30277-1",
				themeId: "65.481.858",
			}, {
				name: "Star Destroyer + TIE Fighter - Mini foil pack",
				setNumber: "911510-1",
				themeId: "65.481.258",
			}, {
				name: "Mini Star Destroyer - Star Wars Celebration Anaheim 2015",
				setNumber: "CELEB2015SD-1",
				themeId: "65.983",
			}, {
				name: "Advent Calendar 2015, Star Wars (Day 11) - Star Destroyer",
				setNumber: "75097-12",
				themeId: "390.715.65",
			},
		];

		let json = readTextFile("./fixtures/searchLegoSets/success.json");
		fetchMock.mockResolvedValue(new Response(json, { status: 200 }));

		let searchResults = await webScrapper.searchLegoSets("destroyer", { themeId: "65", startYear: "2014", endYear: "2017" });

		expect(searchResults.length).toEqual(7);
		expect(searchResults).toEqual(result);
	});

	test("Throws error when not ok", async () => {
		fetchMock.mockResolvedValue(new Response("", { status: 400, statusText: "Bad Request" }));

		await expect(webScrapper.searchLegoSets("")).rejects.toThrow("Failed to fetch data");
	});

	test("Invalid parameters", async () => {
		let json = readTextFile("./fixtures/searchLegoSets/error.json");
		fetchMock.mockResolvedValue(new Response(json, { status: 200 }));

		// let searchResults = await webScrapper.searchLegoSets("");
		await expect(webScrapper.searchLegoSets("")).rejects.toThrow("Query keyword is not specified!");
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

	test("Not Found", async () => {
		let setNumber = "10679-1";

		let html = readTextFile("./fixtures/getLegoSetInfo/not-found.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		await expect(webScrapper.getLegoSetInfo(setNumber)).rejects.toThrow(`Could not find the set ${setNumber}`);
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

	test("Return empty array if no pieces are found", async () => {
		/** @type {LegoSetPieceInfo} */
		let result = {
			normalPieces: [],
			minifigs: [],
			extraPieces: [],
			counterparts: []
		};

		let html = readTextFile("./fixtures/getLegoSetPieces/empty.html");

		fetchMock.mockResolvedValue(new Response(html, { status: 200 }));

		let legoSetPieceInfo = await webScrapper.getLegoSetPieces("10679-1");

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

		await webScrapper.downloadImage(url, downloadFile);

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
		await webScrapper.downloadImage(url, downloadFile);

		// Check if file still exists after attempting to download again
		expect(fetch).not.toHaveBeenCalledWith(url);
		expect(fs.existsSync(downloadFile)).toEqual(true);
	});

	test("Throws error when the response is not ok", async () => {
		fetchMock.mockResolvedValue(new Response(imageData, { status: 404, statusText: "Not Found" }));

		await expect(webScrapper.downloadImage("https://example.com/image.jpg", downloadFile)).rejects.toThrow(
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

		await expect(webScrapper.downloadImage(url, downloadFile)).rejects.toThrow("Write failed");
	});
});
