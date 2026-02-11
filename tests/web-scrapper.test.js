// Imports
import { afterEach, beforeEach, describe, expect, jest, test } from "@jest/globals";
import fs from "fs";
import fsPromises from "fs/promises";
import Webscraper from "../src/web-scrapper.js";
import { getRelativeFilePath, readTextFile } from "./testFunctions.js";


// Variables
/** @type {Color[]} */
let colors;
/** @type {Webscraper} */
let webscraper;

/** @type {jest.MockedFunction<fetch>} */
let fetchMock;


// Setup / Teardown
beforeEach(() => {
	fetchMock = jest.spyOn(global, "fetch");
	colors = [];
	webscraper = new Webscraper(colors);
});

afterEach(() => {
	jest.resetAllMocks();
	jest.restoreAllMocks();
});


// Tests
describe("getWebpage", () => {
	test("Fetches the URL and returns a document on success", async () => {
		let html = readTextFile("./fixtures/getWebpage/success.html");
		let url = "https://example.com";

		// Mock fetch to return a successful response with HTML content
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		let document = await webscraper.getWebpage(url);

		expect(fetch).toHaveBeenCalledWith(url);
		// Verify that the document is a valid HTML document and contains
		expect(document.constructor.name).toEqual("Document");
		expect(document.documentElement.tagName).toEqual("HTML");
		expect(document.querySelector("h1").textContent).toBe("Hello World!");
	});

	test("Throws error when the response is not ok", async () => {
		// Mock fetch to return a failed response
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: false,
			status: 404,
			statusText: "Not Found"
		}));

		// Expect an error to be thrown when the response is
		await expect(webscraper.getWebpage("https://example.com")).rejects.toThrow(
			new Error("Failed to fetch https://example.com: 404 Not Found")
		);
	});
});

describe("getColors", () => {
	test("Returns an array of colors", async () => {
		let html = readTextFile("./fixtures/getColors/colors.html");

		// Mock fetch to return the HTML content of the fixture
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		let colors = await webscraper.getColors();

		// Verify that the returned value is the correct array of colors
		expect(colors).toBeInstanceOf(Array);
		expect(colors.length).toEqual(7);
		expect(colors).toEqual([
			{ bricklinkId: 1, bricklinkName: "White", legoId: 1, legoName: "White" },
			{ bricklinkId: 5, bricklinkName: "Red", legoId: 21, legoName: "Bright Red" },
			{ bricklinkId: 48, bricklinkName: "Sand Green", legoId: 151, legoName: "Sand Green" },
			{ bricklinkId: 12, bricklinkName: "Trans-Clear", legoId: 40, legoName: "Transparent" },
			{ bricklinkId: 17, bricklinkName: "Trans-Red", legoId: 41, legoName: "Tr. Red" },
			{ bricklinkId: 14, bricklinkName: "Trans-Dark Blue", legoId: 43, legoName: "Tr. Blue" },
			{ bricklinkId: 122, bricklinkName: "Chrome Black", legoId: null, legoName: null }
		]);
	});

	test("Throws an error", async () => {
		// Mock the fetch function to simulate a network error
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: false,
			status: 404,
			statusText: "Not Found"
		}));

		// Expect an error to be thrown when calling getColors
		await expect(webscraper.getColors()).rejects.toThrow(
			"Failed to fetch https://v2.bricklink.com/en-us/catalog/color-guide: 404 Not Found"
		);
	});

	test("Returns an empty array", async () => {
		let html = readTextFile("./fixtures/getColors/empty.html");

		// Mock the fetch function to return a response with the
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		let colors = await webscraper.getColors();

		// Verify that the function returns an empty array
		expect(colors).toBeInstanceOf(Array);
		expect(colors.length).toEqual(0);
		expect(colors).toEqual([]);
	});
});

describe("getMinifigPieces", () => {
	// Setup the mock colors before each test
	beforeEach(() => {
		colors.push(
			{ bricklinkId: 11, bricklinkName: "Black", legoId: 26, legoName: "Black" },
			{ bricklinkId: 5, bricklinkName: "Red", legoId: 21, legoName: "Bright Red" },
			{ bricklinkId: 89, bricklinkName: "Dark Purple", legoId: 268, legoName: "Medium Lilac" }
		);
	});

	test("Returns an array of minifig pieces", async () => {
		let html = readTextFile("./fixtures/getMinifigPieces/minifig.html");

		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		/** @type {SetPiece[]} */
		let result = [
			{
				piece: {
					bricklinkId: "970c00",
					name: "Hips and Legs Plain",
					color: {
						bricklinkId: 11,
						bricklinkName: "Black",
						legoId: 26,
						legoName: "Black"
					},
					category: "Minifigure, Legs"
				},
				type: "counterpart",
				amountNeeded: 1
			},
			{
				piece: {
					bricklinkId: "973c000",
					name: "Torso Plain / (Same Color) Arms / (Same Color) Hands",
					color: {
						bricklinkId: 5,
						bricklinkName: "Red",
						legoId: 21,
						legoName: "Bright Red"
					},
					category: "Minifigure, Torso Assembly"
				},
				type: "counterpart",
				amountNeeded: 1
			},
			{
				piece: {
					bricklinkId: "3626",
					name: "Minifigure, Head (Plain)",
					color: {
						bricklinkId: 89,
						bricklinkName: "Dark Purple",
						legoId: 268,
						legoName: "Medium Lilac"
					},
					category: "Minifigure, Head"
				},
				type: "counterpart",
				amountNeeded: 1
			}
		];

		let pieces = await webscraper.getMinifigPieces();

		expect(pieces).toEqual(result);
	});

	test("Returns an empty array of minifig pieces", async () => {
		let html = readTextFile("./fixtures/getMinifigPieces/empty.html");

		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		/** @type {SetPiece[]} */
		let result = [];

		let pieces = await webscraper.getMinifigPieces();

		expect(pieces).toEqual(result);
	});

	test("Missing color", async () => {
		let html = readTextFile("./fixtures/getMinifigPieces/missing-color.html");

		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		/** @type {SetPiece[]} */
		let result = [{
			piece: {
				bricklinkId: "970c00",
				name: "Hips and Legs Plain",
				color: null,
				category: "Minifigure, Legs"
			},
			type: "counterpart",
			amountNeeded: 1
		},];

		let pieces = await webscraper.getMinifigPieces();

		expect(pieces).toEqual(result);
	});

	test("Missing parts sections", async () => {
		let html = readTextFile("./fixtures/getMinifigPieces/missing-parts-section.html");

		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		/** @type {SetPiece[]} */
		let result = [{
			piece: {
				bricklinkId: "970c00",
				name: "Hips and Legs Plain",
				color: {
					bricklinkId: 11,
					bricklinkName: "Black",
					legoId: 26,
					legoName: "Black"
				},
				category: "Minifigure, Legs"
			},
			type: "counterpart",
			amountNeeded: 1
		}];

		expect(await webscraper.getMinifigPieces()).toEqual(result);
	});

	test("Missing regular items section", async () => {
		let html = readTextFile("./fixtures/getMinifigPieces/missing-regular-items-section.html");

		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		/** @type {SetPiece[]} */
		let result = [];

		expect(await webscraper.getMinifigPieces()).toEqual([]);
	});
});

describe("getCompositePiece", () => {
	// TODO: Implement after MVP
	test.todo("Implement after MVP");
});

describe("getLegoSet", () => {
	beforeEach(() => {
		colors.push(
			{
				bricklinkId: 1,
				bricklinkName: "White",
				legoId: 1,
				legoName: "White"
			},
			{
				bricklinkId: 85,
				bricklinkName: "Dark Blueish Gray",
				legoId: 199,
				legoName: "Dark Stone Grey"
			},
			{
				bricklinkId: 11,
				bricklinkName: "Black",
				legoId: 26,
				legoName: "Black"
			},
			{
				bricklinkId: 2,
				bricklinkName: "Tan",
				legoId: 5,
				legoName: "Brick Yellow"
			},
			{
				bricklinkId: 14,
				bricklinkName: "Trans-Dark Blue",
				legoId: 43,
				legoName: "Tr. Blue"
			}
		);
	});

	test("Successfully fetch Lego set", async () => {
		let result = {
			name: "Pirate Treasure Hunt",
			setNumber: "10679-1",
			theme: "Juniors, Pirates, Pirates III",
			year: 2015,
			pieceCount: 46,
			minifigCount: 2,
			normalPieces: [
				{
					piece: {
						bricklinkId: "4738a",
						name: "Container, Treasure Chest Bottom with Slots in Back",
						color: colors[2],
						category: "Container"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "4739a",
						name: "Container, Treasure Chest Lid Curved with Thick Hinge",
						color: colors[2],
						category: "Container"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "92338",
						name: "Chain 5 Links",
						color: colors[1],
						category: "Chain"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "14518",
						name: "Shark Body with Debossed Gills",
						color: colors[1],
						category: "Animal, Body Part"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "87587",
						name: "Shark Head with Rounded Nose and Debossed Eyes",
						color: colors[1],
						category: "Animal, Body Part"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "3068pb0906",
						name: "Tile 2 x 2 with Map Blue Water, Lime Land, Sailing Ship, Treasure Chest and Red 'X' Pattern",
						color: colors[3],
						category: "Tile, Decorated"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "30153",
						name: "Rock 1 x 1 Jewel 24 Facet",
						color: colors[4],
						category: "Rock"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "2335pb129",
						name: "Flag 2 x 2 Square with Skull and Crossbones with No Lower Jaw on Black Background Pattern on Both Sides (Jolly Roger)",
						color: colors[0],
						category: "Flag, Decorated"
					},
					type: "counterpart",
					amountNeeded: 1
				}
			],
			minifigs: [
				{
					piece: {
						bricklinkId: "pi146",
						name: "Pirate Blue Jacket, Black Leg with Peg Leg, Black Pirate Hat with Skull",
						color: null,
						category: "Pirates"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "gen067",
						name: "Skeleton - Standard Skull, Floppy Arms, Red Bandana with Double Tail in Back",
						color: null,
						category: "Pirates"
					},
					type: "counterpart",
					amountNeeded: 1
				}
			],
			extraPieces: [
				{
					piece: {
						bricklinkId: "92338",
						name: "Chain 5 Links",
						color: colors[1],
						category: "Chain"
					},
					type: "counterpart",
					amountNeeded: 1
				}
			],
			counterparts: [
				{
					piece: {
						bricklinkId: "4738ac01",
						name: "Container, Treasure Chest with Slots in Back and (Same Color) Thick Hinge Curved Lid (4738a / 4739a)",
						color: colors[2],
						category: "Container"
					},
					type: "counterpart",
					amountNeeded: 1
				},
				{
					piece: {
						bricklinkId: "14518c01",
						name: "Shark with Rounded Nose and Debossed Gills and Eyes",
						color: colors[1],
						category: "Animal, Water"
					},
					type: "counterpart",
					amountNeeded: 1
				}
			]
		};

		let legoSetInfoHtml = readTextFile("./fixtures/getLegoSetInfo/success.html");
		let legoSetPiecesHtml = readTextFile("./fixtures/getLegoSetPieces/success.html");

		fetchMock
			.mockResolvedValueOnce(/** @type {Response} */({
				ok: true,
				text: () => Promise.resolve(legoSetInfoHtml)
			}))
			.mockResolvedValueOnce(/** @type {Response} */({
				ok: true,
				text: () => Promise.resolve(legoSetPiecesHtml)
			}));

		let legoSet = await webscraper.getLegoSet("10679-1");

		expect(legoSet.normalPieces[0].piece.color).toBe(result.normalPieces[0].piece.color);
		expect(legoSet).toEqual(result);
	});

	test("Not found", async () => {
		let legoSetInfoHtml = readTextFile("./fixtures/getLegoSetInfo/not-found.html");
		let legoSetPiecesHtml = readTextFile("./fixtures/getLegoSetPieces/success.html");

		fetchMock
			.mockResolvedValueOnce(/** @type {Response} */({
				ok: true,
				text: () => Promise.resolve(legoSetInfoHtml)
			}))
			.mockResolvedValueOnce(/** @type {Response} */({
				ok: true,
				text: () => Promise.resolve(legoSetPiecesHtml)
			}));

		await expect(webscraper.getLegoSet("10679-2")).rejects.toThrow("Could not find the set 10679-2");
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
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			arrayBuffer: () => Promise.resolve(imageData)
		}));

		// Verify that the file does not exist before downloading it
		expect(fs.existsSync(downloadFile)).toEqual(false);

		await webscraper.downloadImage(url, downloadFile);

		// Verify that the file was downloaded
		expect(fetch).toHaveBeenCalledWith(url);
		expect(fs.existsSync(downloadFile)).toEqual(true);
		expect(fs.readFileSync(downloadFile)).toEqual(fs.readFileSync(sourceFile));
	});

	test("Does not download image if already exists", async () => {
		fs.writeFileSync(downloadFile, imageData);

		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			arrayBuffer: () => Promise.resolve(imageData)
		}));

		// Check if file exists before downloading
		expect(fs.existsSync(downloadFile)).toEqual(true);
		await webscraper.downloadImage(url, downloadFile);

		// Check if file still exists after attempting to download again
		expect(fetch).not.toHaveBeenCalledWith(url);
		expect(fs.existsSync(downloadFile)).toEqual(true);
	});

	test("Throws error when the response is not ok", async () => {
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: false,
			status: 404,
			statusText: "Not Found"
		}));

		await expect(webscraper.downloadImage("https://example.com/image.jpg", downloadFile)).rejects.toThrow(
			new Error("Failed to fetch https://example.com/image.jpg: 404 Not Found")
		);
	});

	test("Throws error when write fails", async () => {
		fetchMock.mockResolvedValue({
			ok: true,
			arrayBuffer: () => Promise.resolve(imageData),
		});

		// Mock fsPromises.writeFile to simulate a write failure
		let writeFileMock = jest.spyOn(fsPromises, "writeFile");
		writeFileMock.mockImplementation(() => {
			throw new Error("Write failed");
		});

		await expect(webscraper.downloadImage(url, downloadFile)).rejects.toThrow("Write failed");
	});
});
