// Imports
import { afterEach, beforeEach, describe, expect, jest, test } from "@jest/globals";
import fs from "fs";
import fsPromises from "fs/promises";
import Webscraper from "../src/webScrapper.js";
import { getRelativeFilePath, readTextFile } from "./testFunctions.js";
import { LegoColor, LegoPiece, LegoSet, LegoSetPiece } from "../src/classes.js";


// Variables
/** @type {LegoColor[]} */
let colors;
/** @type {LegoPiece[]} */
let legoPieces;
/** @type {Webscraper} */
let webscraper;

/** @type {jest.SpiedFunction<fetch>} */
let fetchMock;


// Setup / Teardown
beforeEach(() => {
	fetchMock = jest.spyOn(global, "fetch");
	colors = [];
	legoPieces = [];
	webscraper = new Webscraper(colors, legoPieces);
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
		fetchMock.mockResolvedValue(/** @type {Response} */({
			ok: true,
			text: () => Promise.resolve(html)
		}));

		let document = await webscraper.getWebpage(url);

		expect(fetch).toHaveBeenCalledWith(url);
		// Verify that the document is a valid HTML document and contains
		expect(document.constructor.name).toEqual("Document");
		expect(document.documentElement.tagName).toEqual("HTML");
		expect(document.querySelector("h1")?.textContent).toBe("Hello World!");
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
		expect(colors[0]).toBeInstanceOf(LegoColor);
		expect(colors).toEqual([
			new LegoColor(null, 1, "White", 1, "White"),
			new LegoColor(null, 5, "Red", 21, "Bright Red"),
			new LegoColor(null, 48, "Sand Green", 151, "Sand Green"),
			new LegoColor(null, 12, "Trans-Clear", 40, "Transparent"),
			new LegoColor(null, 17, "Trans-Red", 41, "Tr. Red"),
			new LegoColor(null, 14, "Trans-Dark Blue", 43, "Tr. Blue"),
			new LegoColor(null, 122, "Chrome Black", null, "")
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
	// TODO: Implement after MVP
	test.todo("Implement after MVP");
	// // Setup the mock colors before each test
	// beforeEach(() => {
	// 	colors.push(
	// 		{ bricklinkId: 11, bricklinkName: "Black", legoId: 26, legoName: "Black" },
	// 		{ bricklinkId: 5, bricklinkName: "Red", legoId: 21, legoName: "Bright Red" },
	// 		{ bricklinkId: 89, bricklinkName: "Dark Purple", legoId: 268, legoName: "Medium Lilac" }
	// 	);
	// });

	// test("Returns an array of minifig pieces", async () => {
	// 	let html = readTextFile("./fixtures/getMinifigPieces/minifig.html");

	// 	fetchMock.mockResolvedValue(/** @type {Response} */({
	// 		ok: true,
	// 		text: () => Promise.resolve(html)
	// 	}));

	// 	/** @type {SetPiece[]} */
	// 	let result = [
	// 		{
	// 			piece: {
	// 				bricklinkId: "970c00",
	// 				name: "Hips and Legs Plain",
	// 				color: {
	// 					bricklinkId: 11,
	// 					bricklinkName: "Black",
	// 					legoId: 26,
	// 					legoName: "Black"
	// 				},
	// 				category: "Minifigure, Legs"
	// 			},
	// 			type: "counterpart",
	// 			amountNeeded: 1
	// 		},
	// 		{
	// 			piece: {
	// 				bricklinkId: "973c000",
	// 				name: "Torso Plain / (Same Color) Arms / (Same Color) Hands",
	// 				color: {
	// 					bricklinkId: 5,
	// 					bricklinkName: "Red",
	// 					legoId: 21,
	// 					legoName: "Bright Red"
	// 				},
	// 				category: "Minifigure, Torso Assembly"
	// 			},
	// 			type: "counterpart",
	// 			amountNeeded: 1
	// 		},
	// 		{
	// 			piece: {
	// 				bricklinkId: "3626",
	// 				name: "Minifigure, Head (Plain)",
	// 				color: {
	// 					bricklinkId: 89,
	// 					bricklinkName: "Dark Purple",
	// 					legoId: 268,
	// 					legoName: "Medium Lilac"
	// 				},
	// 				category: "Minifigure, Head"
	// 			},
	// 			type: "counterpart",
	// 			amountNeeded: 1
	// 		}
	// 	];

	// 	let pieces = await webscraper.getMinifigPieces();

	// 	expect(pieces).toEqual(result);
	// });

	// test("Returns an empty array of minifig pieces", async () => {
	// 	let html = readTextFile("./fixtures/getMinifigPieces/empty.html");

	// 	fetchMock.mockResolvedValue(/** @type {Response} */({
	// 		ok: true,
	// 		text: () => Promise.resolve(html)
	// 	}));

	// 	/** @type {SetPiece[]} */
	// 	let result = [];

	// 	let pieces = await webscraper.getMinifigPieces();

	// 	expect(pieces).toEqual(result);
	// });

	// test("Missing color", async () => {
	// 	let html = readTextFile("./fixtures/getMinifigPieces/missing-color.html");

	// 	fetchMock.mockResolvedValue(/** @type {Response} */({
	// 		ok: true,
	// 		text: () => Promise.resolve(html)
	// 	}));

	// 	/** @type {SetPiece[]} */
	// 	let result = [{
	// 		piece: {
	// 			bricklinkId: "970c00",
	// 			name: "Hips and Legs Plain",
	// 			color: null,
	// 			category: "Minifigure, Legs"
	// 		},
	// 		type: "counterpart",
	// 		amountNeeded: 1
	// 	},];

	// 	let pieces = await webscraper.getMinifigPieces();

	// 	expect(pieces).toEqual(result);
	// });

	// test("Missing parts sections", async () => {
	// 	let html = readTextFile("./fixtures/getMinifigPieces/missing-parts-section.html");

	// 	fetchMock.mockResolvedValue(/** @type {Response} */({
	// 		ok: true,
	// 		text: () => Promise.resolve(html)
	// 	}));

	// 	/** @type {SetPiece[]} */
	// 	let result = [{
	// 		piece: {
	// 			bricklinkId: "970c00",
	// 			name: "Hips and Legs Plain",
	// 			color: {
	// 				bricklinkId: 11,
	// 				bricklinkName: "Black",
	// 				legoId: 26,
	// 				legoName: "Black"
	// 			},
	// 			category: "Minifigure, Legs"
	// 		},
	// 		type: "counterpart",
	// 		amountNeeded: 1
	// 	}];

	// 	expect(await webscraper.getMinifigPieces()).toEqual(result);
	// });

	// test("Missing regular items section", async () => {
	// 	let html = readTextFile("./fixtures/getMinifigPieces/missing-regular-items-section.html");

	// 	fetchMock.mockResolvedValue(/** @type {Response} */({
	// 		ok: true,
	// 		text: () => Promise.resolve(html)
	// 	}));

	// 	/** @type {SetPiece[]} */
	// 	let result = [];

	// 	expect(await webscraper.getMinifigPieces()).toEqual([]);
	// });
});

describe("getCompositePiece", () => {
	// TODO: Implement after MVP
	test.todo("Implement after MVP");
});

describe("getLegoSet", () => {
	beforeEach(() => {
		colors.push(
			new LegoColor(1, 1, "White", 1, "White"),
			new LegoColor(2, 85, "Dark Blueish Gray", 199, "Dark Stone Grey"),
			new LegoColor(3, 11, "Black", 26, "Black"),
			new LegoColor(4, 2, "Tan", 5, "Brick Yellow"),
			new LegoColor(5, 14, "Trans-Dark Blue", 43, "Tr. Blue")
		);
		legoPieces.push(
			new LegoPiece(1, "4738a", "Container, Treasure Chest Bottom with Slots in Back", colors[2], "Container"),
			new LegoPiece(2, "4739a", "Container, Treasure Chest Lid Curved with Thick Hinge", colors[2], "Container"),
			new LegoPiece(3, "92338", "Chain 5 Links", colors[1], "Chain"),
			new LegoPiece(
				4,
				"3068pb0906",
				"Tile 2 x 2 with Map Blue Water, Lime Land, Sailing Ship, Treasure Chest and Red 'X' Pattern",
				colors[3],
				"Tile, Decorated"
			),
			new LegoPiece(5, "pi146", "Pirate Blue Jacket, Black Leg with Peg Leg, Black Pirate Hat with Skull", null, "Pirates"),
			new LegoPiece(6, "gen067", "Skeleton - Standard Skull, Floppy Arms, Red Bandana with Double Tail in Back", null, "Pirates"),
			new LegoPiece(7, "92338", "Chain 5 Links", colors[1], "Chain"),
			new LegoPiece(
				8,
				"4738ac01",
				"Container, Treasure Chest with Slots in Back and (Same Color) Thick Hinge Curved Lid (4738a / 4739a)",
				colors[2],
				"Container"
			)
		);
	});

	test("Successfully fetch Lego set", async () => {
		let result = new LegoSet(
			null,
			"10679-1",
			"Pirate Treasure Hunt",
			"Juniors, Pirates, Pirates III",
			2015,
			46,
			2,
			1,
			[
				new LegoSetPiece(null, legoPieces[0], 1),
				new LegoSetPiece(null, legoPieces[1], 1),
				new LegoSetPiece(null, legoPieces[2], 1),
				new LegoSetPiece(null, legoPieces[3], 10),
			],
			[
				new LegoSetPiece(null, legoPieces[4], 1),
				new LegoSetPiece(null, legoPieces[5], 1)
			],
			[
				new LegoSetPiece(null, legoPieces[6], 1)
			],
			[
				new LegoSetPiece(null, legoPieces[7], 1)
			]
		);

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

		expect(legoSet.normalPieces[0].color).toBe(result.normalPieces[0].color);
		expect(legoSet.normalPieces[0]).toEqual(result.normalPieces[0]);
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
		// fetchMock.mockResolvedValue(/** @type {Response} */({
		// 	ok: true,
		// 	arrayBuffer: () => Promise.resolve(imageData)
		// }));
		// fetchMock.mockResolvedValue(new Response({ arrayBuffer: () => Promise.resolve(imageData) }, { status: 200 }));
		fetchMock.mockResolvedValue(new Response(imageData, { status: 200 }));

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

		fetchMock.mockResolvedValue(new Response(imageData, { status: 200 }));

		// Check if file exists before downloading
		expect(fs.existsSync(downloadFile)).toEqual(true);
		await webscraper.downloadImage(url, downloadFile);

		// Check if file still exists after attempting to download again
		expect(fetch).not.toHaveBeenCalledWith(url);
		expect(fs.existsSync(downloadFile)).toEqual(true);
	});

	test("Throws error when the response is not ok", async () => {
		fetchMock.mockResolvedValue(new Response(imageData, { status: 404, statusText: "Not Found" }));

		await expect(webscraper.downloadImage("https://example.com/image.jpg", downloadFile)).rejects.toThrow(
			new Error("Failed to fetch https://example.com/image.jpg: 404 Not Found")
		);
	});

	test("Throws error when write fails", async () => {
		fetchMock.mockResolvedValue(new Response(imageData, { status: 200 }));

		// Mock fsPromises.writeFile to simulate a write failure
		let writeFileMock = jest.spyOn(fsPromises, "writeFile");
		writeFileMock.mockImplementation(() => {
			throw new Error("Write failed");
		});

		await expect(webscraper.downloadImage(url, downloadFile)).rejects.toThrow("Write failed");
	});
});
