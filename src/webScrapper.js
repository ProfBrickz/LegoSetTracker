// Imports
import fs from "fs";
import fsPromises from "fs/promises";
import { JSDOM } from "jsdom";
import path from "path";
import { MINIFIG_IMAGES_PATH, PIECE_IMAGES_PATH, SET_IMAGES_PATH } from "./constants.js";
import { LegoColors } from "./controllers.js";
import { LegoColor, LegoPiece, LegoSet, LegoSetPiece } from "./models.js";
/** @import {LegoSetInfo, LegoSetPieceInfo, LegoSetSearchResult} from "./types.js" */


// Functions
export default class WebScrapper {
	/** @type {LegoColors} */
	#colors;

	/**
	 * @public
	 * @param {LegoColors} colors A map of Lego colors
	 */
	constructor(colors) {
		this.#colors = colors;
	}

	/**
	 * Fetches and parses the HTML content of a webpage from the given URL, returning the parsed Document object.
	 *
	 * @public
	 * @param {string} url The URL of the webpage to fetch and parse.
	 * @returns {Promise<Document>} A Promise that resolves to the parsed Document object once the request is successful.
	 * @throws {Error} If the request fails (e.g., network error, HTTP error status code).
	 */
	async getWebpage(url) {
		let response = await fetch(url);

		if (!response.ok) {
			throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
		}

		let html = await response.text();
		let { document } = new JSDOM(html).window;
		return document;
	}

	/**
	 * Fetches and parses LEGO set information from BrickLink's website.
	 *
	 * This function retrieves detailed information about a LEGO set (name, number, theme, year,
	 * piece count, and minifigure count) by querying BrickLink's catalog and parsing the HTML response.
	 *
	 * @public
	 * @param {string} setNumberInput The LEGO set number (e.g., "8699-1") used to construct the query URL.
	 * @returns {Promise<LegoSetInfo>} A Promise that resolves to an object containing the parsed set information.
	 * @throws {Error} If the set number is invalid, the document cannot be parsed, or required elements are missing.
	 */
	async getLegoSetInfo(setNumberInput) {
		let document = await this.getWebpage(`https://www.bricklink.com/v2/catalog/catalogitem.page?S=${setNumberInput}`);

		// Extracting the name from the document
		let nameElement = /** @type {HTMLHeadingElement} */ (document.querySelector("#item-name-title"));
		if (!nameElement) {
			throw new Error(`Could not find the set ${setNumberInput}`);
		}
		let name = nameElement.textContent.trim();


		// Extracting the theme from the document
		let themeElement = /** @type {HTMLTableCellElement} */ (document.querySelector(
			"#content .innercontent table:first-of-type tr td:nth-child(1)"
		));

		let theme = Array.from(themeElement.children)
			.slice(2)
			.map(element => element.textContent)
			.join(", ");

		// Extracting the set number from the document
		let setNumberElement = /** @type {HTMLSpanElement} */ (document.querySelector(
			"#id_divBlock_Main table:first-of-type tr:first-of-type td:first-of-type span span"
		));
		let setNumber = setNumberElement.textContent.trim();

		// Extracting the year from the document
		let yearElement = /** @type {HTMLAnchorElement} */ (document.querySelector(
			"#id_divBlock_Main table:first-of-type tr:nth-of-type(2) td:nth-of-type(2) table tr:first-of-type td:first-of-type a:first-of-type"
		));
		let releaseYear = Number.parseInt(yearElement.textContent.trim());

		// Extracting the piece count from the document
		let pieceCount = 0;
		let minifigCount = 0;
		let pieceAndMinifigLinks = /** @type {NodeListOf<HTMLAnchorElement>} */ (document.querySelectorAll("#id_divBlock_Main table:first-of-type tr:nth-of-type(2) td:nth-of-type(2) table tr:first-of-type td:first-of-type table tr:first-of-type td:nth-of-type(2) a"));
		for (let link of pieceAndMinifigLinks) {
			let text = link.textContent;

			if (text.includes("Minifigure")) {
				minifigCount = Number.parseInt(text.split(" ")[0]);
			} else if (text.includes("Parts")) {
				pieceCount = Number.parseInt(text.split(" ")[0]);
			}
		}

		return { name, setNumber, theme, releaseYear, pieceCount, minifigCount };
	}

	/**
	 * Fetches and parses Lego set pieces information from BrickLink's catalog.
	 *
	 * This function retrieves detailed information about all items in a LEGO set,
	 * categorized into "Regular Items", "Minifigures", "Extra Items", and "Counterparts".
	 * It parses the HTML table structure from BrickLink's inventory page to extract
	 * the relevant data for each category.
	 *
	 * @public
	 * @param {string} setNumber The LEGO set number (e.g., "8699-1") used to construct the query URL.
	 * @returns {Promise<LegoSetPieceInfo>} A Promise that resolves to an object containing categorized piece data.
	 * @throws {Error} If the set number is invalid, the document cannot be parsed, or required elements are missing.
	 */
	async getLegoSetPieces(setNumber) {
		let document = await this.getWebpage(`https://www.bricklink.com/catalogItemInv.asp?S=${setNumber}&viewType=P&sortBy=0&bt=0&sortAsc=a`);
		let tbody = /** @type {HTMLTableSectionElement}*/ (document.querySelector("form > table tbody"));
		let rows = /** @type {HTMLTableRowElement[]} */ (Array.from(tbody.children));
		let categories = /** @type {NodeListOf<HTMLTableRowElement>} */ (tbody.querySelectorAll("tr[bgcolor='#000000'], tr[bgcolor='#C0C0C0']"));

		let normalPieces = this.getSection("Regular Items:", rows, categories);
		let minifigs = this.getSection("Minifigures:", rows, categories);
		let extraPieces = this.getSection("Extra Items:", rows, categories);
		let counterparts = this.getSection("Counterparts:", rows, categories);

		return { normalPieces, minifigs, extraPieces, counterparts };
	}

	/**
	 * Extracts rows corresponding to a specific section from a table.
	 *
	 * This function identifies the rows in a table that belong to a specific section
	 * (e.g., "Regular Items", "Minifigures") by matching the section name in the category rows.
	 * It then slices the relevant rows and processes them to extract piece data.
	 *
	 * @private
	 * @param {string} section The name of the section to extract (e.g., "Regular Items:").
	 * @param {HTMLTableRowElement[]} rows An array of table rows to search through.
	 * @param {NodeListOf<HTMLTableRowElement>} categories A list of category rows used to identify section boundaries.
	 * @returns {LegoSetPiece[]} An array of piece data objects corresponding to the specified section.
	 * @throws {Error} If the specified section cannot be found in the categories.
	 */
	getSection(section, rows, categories) {
		/** @type {number|null} */
		let sectionStartIndex = null;
		/** @type {number|null} */
		let sectionEndIndex = null;

		for (let category of categories) {
			let index = rows.indexOf(category);

			if (category.textContent.trim() == section) {
				sectionStartIndex = index + 1;
				continue;
			}

			if (sectionStartIndex) {
				if (category.textContent.trim() == "Parts:") {
					sectionStartIndex = index + 1;
					continue;
				}

				sectionEndIndex = index;
				break;
			}
		}

		if (sectionStartIndex && !sectionEndIndex) sectionEndIndex = rows.length;
		if (!sectionStartIndex || !sectionEndIndex) return [];

		let categoryRows = rows.slice(sectionStartIndex, sectionEndIndex);

		return this.getSectionPieces(categoryRows);
	}

	/**
	 * Extracts piece data from an array of table rows.
	 *
	 * This function processes each row to extract detailed information about a LEGO piece,
	 * including its bricklink ID, image URL, name, color, category, quantity needed, and
	 * color mapping from a predefined color list. It constructs and returns an array of
	 * `SetPiece` objects representing the extracted data.
	 *
	 * @private
	 * @param {HTMLTableRowElement[]} sectionRows An array of table rows containing piece details.
	 * @returns {LegoSetPiece[]} An array of `SetPiece` objects representing the extracted piece data.
	 */
	getSectionPieces(sectionRows) {
		/** @type {LegoSetPiece[]} */
		let legoSetPieces = [];

		for (let row of sectionRows) {
			let bricklinkIdElement = /** @type {HTMLAnchorElement} */ (row.querySelector("td:nth-of-type(3) a"));
			let bricklinkColorId = Number((new URL(bricklinkIdElement.href, "https://bricklink.com")).searchParams.get("idColor"));
			let bricklinkId = bricklinkIdElement.textContent.trim();
			let nameElement = /** @type {HTMLElement} */ (row.querySelector("td:nth-of-type(4) b"));
			let nameAndColor = nameElement.textContent.trim();
			// Remove repeated spaces
			nameAndColor = nameAndColor
				.replace(/\s+/g, " ")
				.replace("Bluish", "Blueish");

			let categoryElement = /** @type {HTMLTableCellElement} */ (row.querySelector("td:nth-of-type(4) a:nth-of-type(3)"));
			let bricklinkCategory = categoryElement.textContent.trim();

			let amountNeededElement = /** @type {HTMLTableCellElement} */ (row.querySelector("td:nth-of-type(2)"));
			let amountNeeded = Number.parseInt(amountNeededElement.textContent);

			let color = this.#colors.get(bricklinkColorId) || null;
			let bricklinkName = nameAndColor;
			if (color) bricklinkName = nameAndColor.replace(color.bricklinkName, "").trim();

			legoSetPieces.push(new LegoSetPiece(
				null,
				new LegoPiece(null, bricklinkId, bricklinkName, color, bricklinkCategory),
				amountNeeded
			));
		}

		return legoSetPieces;
	}

	/**
	 * Fetches and processes color data from BrickLink's color guide page.
	 *
	 * This function retrieves the color guide page from BrickLink's website, parses the HTML
	 * to extract color information, and returns an array of color objects containing
	 * BrickLink and LEGO-specific identifiers and names. It relies on a specific HTML structure
	 * for element selection.
	 *
	 * @public
	 * @returns {Promise<LegoColors>} A promise that resolves to an array of `LegoColor` objects.
	 * @throws {Error} If the DOM structure is invalid or required elements are missing.
	 */
	async getColors() {
		let colors = new LegoColors();

		let document = await this.getWebpage("https://v2.bricklink.com/en-us/catalog/color-guide");

		let sections = /** @type {NodeListOf<HTMLTableSectionElement>} */ (document.querySelectorAll("div table:nth-of-type(2) tbody"));

		for (let section of sections) {
			for (let tr of section.children) {
				let bricklinkNameElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(2) p"));
				let bricklinkName = bricklinkNameElement.textContent.trim();

				let legoNameElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(2) span"));

				let [legoName, legoIdStr] = legoNameElement.textContent
					.replace("LEGO", "")
					.replace("Color: ", "")
					.trim()
					.split(" - ");
				/** @type {number | null} */
				let legoId = Number.parseInt(legoIdStr);
				if (!Number.isInteger(legoId)) legoId = null;

				let bricklinkIdElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(8)"));
				let bricklinkId = Number.parseInt(bricklinkIdElement.textContent);

				colors.set(bricklinkId, new LegoColor(null, bricklinkId, bricklinkName, legoId, legoName));
			}
		}

		return colors;
	}

	/**
	 * @param {LegoColors} colors A map of Lego colors
	 */
	setColors(colors) {
		this.#colors = colors;
	}

	/**
	 * Fetches the categories for LEGO sets.
	 *
	 * @public
	 * @returns {Promise<Map<string, string>>}
	 * A promise that resolves to a map where the keys are category names and the values are category IDs.
	 */
	async getLegoSetThemes() {
		/** @type {Map<string, string>} */
		let categories = new Map();

		let document = await this.getWebpage("https://www.bricklink.com/catalogTree.asp?itemType=S");

		let sections =/** @type {HTMLTableRowElement[]} */ Array.from(document.querySelectorAll(".catalog-tree__spacing-reset")).slice(0, 2);
		/** @type {HTMLAnchorElement[]} */
		let links = [];
		for (let section of sections) {
			links.push(...section.querySelectorAll("a"));
		}

		for (let link of links) {
			let id = new URL(link.href, "https://bricklink.com").searchParams.get("catString") || "";
			let ids = id.split(".");
			id = ids[ids.length - 1];

			let name = link.textContent.trim();

			if (name == "{}" || name == "{more}") continue;

			categories.set(id, name);
		}

		return categories;
	}

	/**
	 * @param {string} searchQuery The search query for the search.
	 * @param {Object} [options]
	 * @param {string} [options.themeId] The ID of the theme to filter by.
	 * @param {number} [options.startYear] The start year for the search (inclusive).
	 * @param {number} [options.endYear] The end year for the search (inclusive).
	 *
	 * @returns {Promise<LegoSetSearchResult[]>}
	 */
	async searchLegoSets(searchQuery, { themeId, startYear, endYear } = {}) {
		/** @type {LegoSetSearchResult[]} */
		let legoSets = [];

		let url = new URL("https://www.bricklink.com/ajax/clone/search/searchproduct.ajax");
		url.searchParams.set("type", "S");
		url.searchParams.set("q", searchQuery);
		if (themeId) url.searchParams.set("cat", themeId);
		if (startYear) url.searchParams.set("yf", startYear.toString());
		if (endYear) url.searchParams.set("yt", endYear.toString());

		let response = await fetch(url);
		if (!response.ok) throw new Error("Failed to fetch data");

		/**
		 * @typedef {"S" | "P" | "M" | "G" | "B"} ItemType - The type of the item.
		 *
		 * @typedef {Object} Item
		 * @property {ItemType} typeItem - The type of the item.
		 * @property {string} strItemNo - The item number.
		 * @property {string} strItemName - The name of the item.
		 * @property {string} strCategory - The category of the item.
		 *
		 * @typedef {Object} TypeList
		 * @property {number} type - The type identifier.
		 * @property {number} count - The count of items in this type.
		 * @property {Item[]} items - An array of items belonging to this type.
		 *
		 * @typedef {Object} Result
		 * @property {TypeList[]} typeList - An array of type lists containing items.
		 *
		 * @typedef {Object} SearchResponse
		 * @property {Result} result - The result object containing the data.
		 * @property {number} returnCode - Return code indicating success or failure.
		 * @property {string} returnMessage - Message describing the result.
		 */
		/** @type {SearchResponse} */
		let { result, returnCode, returnMessage } = await response.json();
		if (returnCode != 0) throw new Error(returnMessage);

		let items = result.typeList[0].items;
		items[0].strItemNo;

		for (let item of items) {
			legoSets.push({
				setNumber: item.strItemNo,
				name: item.strItemName,
				themeId: item.strCategory,
			});
		}

		return legoSets;
	}

	/**
	 * Fetches minifig pieces data from BrickLink's catalog for a given minifig ID.
	 *
	 * @public
	 * @param {string} minifigId The BrickLink ID of the minifig (e.g., "3523").
	 * @returns {Promise<LegoSetPiece[]>} A promise that resolves to an array of `SetPiece` objects representing the minifig's parts.
	 */
	async getMinifigPieces(minifigId) {
		throw new Error("TODO");

		// let document = await this.getWebpage(`https://www.bricklink.com/catalogItemInv.asp?M=${minifigId}&viewType=P&bt=0&sortBy=0&sortAsc=a`);
		// let tbody = /** @type {HTMLTableSectionElement}*/ (document.querySelector("form > table tbody"));
		// let rows = /** @type {HTMLTableRowElement[]} */ (Array.from(tbody.children));
		// let categories = /** @type {NodeListOf<HTMLTableRowElement>} */ (tbody.querySelectorAll("tr[bgcolor='#000000'], form > table tr[bgcolor='#C0C0C0']"));

		// return this.getSection("Regular Items:", rows, categories);
	}

	/**
	 * Fetches composite piece data for a specific LEGO part in a given color from BrickLink's catalog.
	 *
	 * @todo
	 * @public
	 * @param {LegoPiece} piece The piece object containing the `bricklinkId` and `color` properties.
	 * @returns {Promise<LegoPiece>} A promise that resolves to a `SetPiece` object representing the piece in the specified color.
	 */
	async getCompositePiece(piece) {
		throw new Error("TODO");

		// let document = await this.getWebpage(`https://www.bricklink.com/catalogItemInv.asp?P=${piece.bricklinkId}&C=${piece.color.bricklinkId}&viewType=P&bt=0&sortBy=0&sortAsc=a`);
		// let tbody = /** @type {HTMLTableSectionElement}*/ (document.querySelector("form > table tbody"));
		// let rows = /** @type {HTMLTableRowElement[]} */ (Array.from(tbody.childNodes));
		// let categories = /** @type {NodeListOf<HTMLTableRowElement>} */ (document.querySelectorAll("form > table tr[bgcolor='#000000'], form > table tr[bgcolor='#C0C0C0']"));

		// let setPieces = this.getSection("Regular Items:", rows, categories);
	}

	/**
	 * Downloads an image from the specified URL to the given download path.
	 *
	 * @public
	 * @param {string} url The URL of the image to download.
	 * @param {string} downloadPath The file path where the image will be saved.
	 * @returns {Promise<void>} A promise that resolves when the image is successfully downloaded or if the target path already exists.
	 * @throws {Error} If the image cannot be fetched (e.g., invalid URL, server error).
	 */
	async downloadImage(url, downloadPath) {
		if (fs.existsSync(downloadPath)) return;

		let response = await fetch(url);

		if (!response.ok) {
			return;
			// throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
		}

		let imageData = await response.arrayBuffer();

		await fsPromises.writeFile(downloadPath, Buffer.from(imageData));
	}

	/**
	 * Downloads the set image from BrickLink for the specified set number.
	 *
	 * @public
	 * @param {string} setNumber The set number (e.g., "10179-1").
	 * @returns {Promise<void>} A promise that resolves when the image is downloaded or if the target path already exists.
	 * @throws {Error} If the image cannot be fetched (e.g., invalid URL, server error).
	 */
	async downloadLegoSetImage(setNumber) {
		// await this.downloadImage(`https://img.bricklink.com/S/${setNumber}.jpg`, path.join(SET_IMAGES_PATH, `${setNumber}.jpg`));
		await this.downloadImage(`https://img.bricklink.com/ItemImage/SN/0/${setNumber}.png`, path.join(SET_IMAGES_PATH, `${setNumber}.jpg`));
	}

	/**
	 * Downloads the piece image from BrickLink for the specified piece ID and color ID.
	 *
	 * @public
	 * @param {string} pieceId The piece ID (e.g., "3003").
	 * @param {number} colorId The color ID (e.g., 11 for red).
	 * @returns {Promise<void>} A promise that resolves when the image is downloaded or if the target path already exists.
	 * @throws {Error} If the image cannot be fetched (e.g., invalid URL, server error) or if writing to the file fails.
	 */
	async downloadLegoPieceImage(pieceId, colorId) {
		await this.downloadImage(`https://img.bricklink.com/P/${colorId}/${pieceId}.jpg`, path.join(PIECE_IMAGES_PATH, colorId.toString(), `${pieceId}.jpg`));
	}

	/**
	 * Downloads the minifig image from BrickLink for the specified minifig ID.
	 *
	 * @public
	 * @param {string} minifigId The minifig ID (e.g., "10179-1").
	 * @returns {Promise<void>} A promise that resolves when the image is downloaded or if the target path already exists.
	 * @throws {Error} If the image cannot be fetched (e.g., invalid URL, server error) or if writing to the file fails.
	 */
	async downloadMinifigImage(minifigId) {
		await this.downloadImage(`https://img.bricklink.com/M/${minifigId}.jpg`, path.join(MINIFIG_IMAGES_PATH, `${minifigId}.jpg`));
	}

	/**
	 * Downloads all images associated with a LEGO set, including set image, pieces, and minifigs.
	 *
	 * @public
	 * @param {LegoSet} legoSet The LEGO set object containing set details, pieces, and minifigs.
	 * @returns {Promise<void>} A promise that resolves when all images are downloaded or if no new images are needed.
	 * @throws {Error} If any of the individual image download operations fail (e.g., network errors, invalid URLs, file write failures).
	 */
	async downloadLegoSetImages(legoSet) {
		await Promise.all([
			this.downloadLegoSetImage(legoSet.setNumber),
			legoSet.normalPieces.forEach((setPiece) =>
				this.downloadLegoPieceImage(setPiece.bricklinkId, setPiece.color?.bricklinkId || 0)
			),
			legoSet.normalPieces.forEach((setPiece) =>
				this.downloadLegoPieceImage(setPiece.bricklinkId, setPiece.color?.bricklinkId || 0)
			),
			legoSet.counterpartPieces.forEach((setPiece) =>
				this.downloadLegoPieceImage(setPiece.bricklinkId, setPiece.color?.bricklinkId || 0)
			),
			legoSet.minifigs.forEach((setPiece) => this.downloadMinifigImage(setPiece.bricklinkId))
		]);
	}
}
