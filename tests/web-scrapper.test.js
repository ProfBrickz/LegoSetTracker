// Imports
import { afterEach, beforeEach, describe, expect, jest, test } from "@jest/globals";
import fs from "fs";
import fsPromises from "fs/promises";
import path from "path";
import { TEST_DIRECTORY } from "../src/constants.js";
import Webscraper from "../src/web-scrapper.js";
import { readTextFile, readBufferFile, getRelativeFilePath } from "./testFunctions.js";


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
	test("fetches the URL and returns a document on success", async () => {
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

	test("throws error when the response is not ok", async () => {
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
	test("returns an array of colors", async () => {
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

	test("throws an error", async () => {
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

	test("returns an empty array", async () => {
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

	test("returns an array of minifig pieces", async () => {
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
					category: "Minifigure, Legs",
					parentId: null,
					children: []
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
					category: "Minifigure, Torso Assembly",
					parentId: null,
					children: []
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
					category: "Minifigure, Head",
					parentId: null,
					children: []
				},
				type: "counterpart",
				amountNeeded: 1
			}
		];

		let pieces = await webscraper.getMinifigPieces();

		expect(pieces).toEqual(result);
	});

	test("returns an empty array of minifig pieces", async () => {
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

	test("missing color", async () => {
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
				category: "Minifigure, Legs",
				parentId: null,
				children: []
			},
			type: "counterpart",
			amountNeeded: 1
		},];

		let pieces = await webscraper.getMinifigPieces();

		expect(pieces).toEqual(result);
	});

	test("missing parts sections", async () => {
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
				category: "Minifigure, Legs",
				parentId: null,
				children: []
			},
			type: "counterpart",
			amountNeeded: 1
		}];

		expect(await webscraper.getMinifigPieces()).toEqual(result);
	});

	test("missing regular items section", async () => {
		let html = readTextFile("./fixtures/getMinifigPieces/missing-regular-items-section.html");

		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		/** @type {SetPiece[]} */
		let result = [];

		await expect(webscraper.getMinifigPieces()).rejects.toThrow("Could not find Regular Items: in categories");
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

	test("downloads image", async () => {
		console.log(fs.readdirSync("."));

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

	test("does not download image if already exists", async () => {
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

	test("throws error when the response is not ok", async () => {
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: false,
			status: 404,
			statusText: "Not Found"
		}));

		await expect(webscraper.downloadImage("https://example.com/image.jpg", downloadFile)).rejects.toThrow(
			new Error("Failed to fetch https://example.com/image.jpg: 404 Not Found")
		);
	});

	test("throws error when write fails", async () => {
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
