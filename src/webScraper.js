// Imports
import fs from "fs";
import fsPromises from "fs/promises";
import { JSDOM } from "jsdom";
import path from "path";
import puppeteer, { Browser, Page } from "puppeteer";
import { MINIFIG_IMAGES_PATH, PIECE_IMAGES_PATH, SET_IMAGES_PATH } from "./constants.js";
import { LegoColors } from "./dataMaps.js";
import { LegoPiece, LegoSet } from "./models.js";
/** @import { LegoSetInfo, LegoSetPieceInfo, LegoSetPiecesInfo, LegoSetSearchResult, WebLegoColor, WebLegoSetTheme } from "./types.js" */


// Functions
export default class WebScraper {
	/**
	 * @private
	 * @type {LegoColors}
	 */
	colors;
	/**
	 * @private
	 * @type {Browser | null}
	 */
	browser = null;
	/**
	 * @private
	 * @type {Page | null}
	 */
	page = null;

	/**
	 * @public
	 * @param {LegoColors} colors A map of Lego colors
	 */
	constructor(colors) {
		this.colors = colors;
	}

	async init() {
		this.browser = await puppeteer.launch({ headless: true });
		this.page = await this.browser.newPage();

		await this.getWebpage("https://www.bricklink.com/v2/main.page");
	}

	async close() {
		if (this.browser == null) throw new Error("Error: Browser is not initialized");

		await this.browser.close();
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
		if (this.browser == null || this.page == null) throw new Error("Error: Browser and page are not initialized");

		await this.page.goto(url, { waitUntil: "networkidle2" });
		const html = await this.page.content();

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

		let themeElements = Array.from(/** @type {HTMLCollectionOf<HTMLLinkElement>} */(themeElement.children));
		let themeId = this.parseThemeIdFromLink(themeElements[themeElements.length - 1].href);

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

		return { name, setNumber, themeId, releaseYear, pieceCount, minifigCount };
	}

	/**
	 * Fetches and parses LEGO set inventory data from BrickLink's website.
	 *
	 * This function retrieves detailed inventory information about items in a LEGO set,
	 * categorized into "Regular Items", "Minifigures", "Extra Items", and "Counterparts".
	 * It parses the HTML table structure from BrickLink's inventory page to extract
	 * the relevant piece data for each category.
	 *
	 * @public
	 * @param {string} url The URL of the BrickLink inventory page (e.g., catalog item inventory URL).
	 * @returns {Promise<{normalPieces: LegoSetPieceInfo[], minifigs: LegoSetPieceInfo[], extraPieces: LegoSetPieceInfo[], counterparts: LegoSetPieceInfo[]}>} A Promise that resolves to an object containing categorized piece data.
	 * @throws {Error} If the webpage cannot be fetched or required elements are missing.
	 */
	async getCatalogItemInventory(url) {
		let document = await this.getWebpage(url);
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
	 * Fetches and parses Lego set pieces information from BrickLink's catalog.a
	 *
	 * @public
	 * @param {string} setNumber The LEGO set number (e.g., "8699-1") used to construct the query URL.
	 * @returns {Promise<LegoSetPiecesInfo>} A Promise that resolves to an object containing categorized piece data.
	 * @throws {Error} If the set number is invalid, the document cannot be parsed, or required elements are missing.
	 */
	async getLegoSetPieces(setNumber) {
		return await this.getCatalogItemInventory(
			`https://www.bricklink.com/catalogItemInv.asp?S=${setNumber}&viewType=P&sortBy=0&sortAsc=A&bt=0`
		);
	}

	/**
 * Fetches minifig pieces data from BrickLink's catalog for a given minifig ID.
 *
 * @public
 * @param {LegoPiece} piece The BrickLink ID of the minifig (e.g., "3523").
 * @returns {Promise<LegoSetPiecesInfo>} A promise that resolves to an array of `SetPiece` objects representing the minifig's parts.
 */
	async getMinifigPieces(piece) {
		return await this.getCatalogItemInventory(
			`https://www.bricklink.com/catalogItemInv.asp?M=${piece.brickLinkId}&viewType=P&sortBy=0&sortAsc=A&bt=0`
		);
	}

	/**
	 * Fetches composite piece data for a specific LEGO part in a given color from BrickLink's catalog.
	 *
	 * @todo
	 * @public
	 * @param {LegoPiece} piece The piece object containing the `brickLinkId` and `color` properties.
	 * @returns {Promise<LegoSetPiecesInfo>} A promise that resolves to a `SetPiece` object representing the piece in the specified color.
	 */
	async getCompositePieceComponents(piece) {
		return await this.getCatalogItemInventory(
			`https://www.bricklink.com/catalogItemInv.asp?P=${piece.brickLinkId}&viewType=P&sortBy=0&sortAsc=A&bt=0`
		);
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
	 * @returns {LegoSetPieceInfo[]} An array of piece data objects corresponding to the specified section.
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
	 * including its BrickLink ID, image URL, name, color, category, quantity needed, and
	 * color mapping from a predefined color list. It constructs and returns an array of
	 * `SetPiece` objects representing the extracted data.
	 *
	 * @private
	 * @param {HTMLTableRowElement[]} sectionRows An array of table rows containing piece details.
	 * @returns {LegoSetPieceInfo[]} An array of `SetPiece` objects representing the extracted piece data.
	 */
	getSectionPieces(sectionRows) {
		/** @type {LegoSetPieceInfo[]} */
		let legoSetPieces = [];

		for (let row of sectionRows) {
			let brickLinkIdElement = /** @type {HTMLAnchorElement} */ (row.querySelector("td:nth-of-type(3) a"));
			let brickLinkColorId = Number((new URL(brickLinkIdElement.href, "https://bricklink.com")).searchParams.get("idColor"));
			let brickLinkId = brickLinkIdElement.textContent.trim();
			let nameElement = /** @type {HTMLElement} */ (row.querySelector("td:nth-of-type(4) b"));
			let nameAndColor = nameElement.textContent.trim();
			// Remove repeated spaces
			nameAndColor = nameAndColor
				.replace(/\s+/g, " ");

			let categoryElement = /** @type {HTMLTableCellElement} */ (row.querySelector("td:nth-of-type(4) a:nth-of-type(3)"));
			let brickLinkCategory = categoryElement.textContent.trim();

			let amountNeededElement = /** @type {HTMLTableCellElement} */ (row.querySelector("td:nth-of-type(2)"));
			let amountNeeded = Number.parseInt(amountNeededElement.textContent);

			let color = this.colors.getByBricklinkId(brickLinkColorId);
			let brickLinkName = nameAndColor;
			if (color) brickLinkName = nameAndColor.replace(color.brickLinkName, "").trim();

			let inventoryLinkElement = /** @type {HTMLAnchorElement | null} */ (row.querySelector("td:nth-of-type(3) a:nth-of-type(2)"));
			let isCompoundPiece = inventoryLinkElement !== null && inventoryLinkElement.textContent === "Inv";

			legoSetPieces.push({
				brickLinkId,
				brickLinkName,
				color,
				brickLinkCategory,
				amountNeeded,
				amountFound: 0,
				isCompoundPiece
			});
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
	 * @returns {Promise<WebLegoColor[]>} A promise that resolves to an array of `LegoColor` objects.
	 * @throws {Error} If the DOM structure is invalid or required elements are missing.
	 */
	async getColors() {
		let colors = [];

		let document = await this.getWebpage("https://v2.bricklink.com/en-us/catalog/color-guide");

		let sections = /** @type {NodeListOf<HTMLTableSectionElement>} */ (document.querySelectorAll("div table:nth-of-type(2) tbody"));

		for (let section of sections) {
			for (let tr of section.children) {
				let brickLinkNameElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(2) p"));
				let brickLinkName = brickLinkNameElement.textContent.trim();

				let legoNameElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(2) span"));

				/** @type {(string | null)[]} */
				let [legoName, legoIdStr] = legoNameElement.textContent
					.replace("LEGO", "")
					.replace("Color: ", "")
					.trim()
					.split(" - ");
				/** @type {number | null} */
				let legoId = Number.parseInt(legoIdStr);
				if (!Number.isInteger(legoId)) legoId = null;
				if (!legoName) legoName = null;

				let brickLinkIdElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(8)"));
				let brickLinkId = Number.parseInt(brickLinkIdElement.textContent);

				colors.push({ brickLinkId, brickLinkName, legoId, legoName });
			}
		}

		return colors;
	}

	/**
	 * Fetches the categories for LEGO sets.
	 */
	async getLegoSetThemes() {
		/** @type {WebLegoSetTheme[]} */
		let themes = [];

		let document = await this.getWebpage("https://www.bricklink.com/catalogTree.asp?itemType=S");

		let sections =/** @type {HTMLTableRowElement[]} */ Array.from(document.querySelectorAll(".catalog-tree__spacing-reset")).slice(0, 2);
		/** @type {HTMLAnchorElement[]} */
		let links = [];
		for (let section of sections) {
			links.push(...section.querySelectorAll("a"));
		}

		for (let link of links) {
			let id = this.parseThemeIdFromLink(link.href);

			let name = link.textContent.trim();

			if (name == "{}" || name == "{more}") continue;

			themes.push({
				brickLinkId: id,
				brickLinkName: name
			});
		}

		return themes;
	}

	/**
	 * @private
	 * @param {string} link
	 */
	parseThemeIdFromLink(link) {
		let id = new URL(link, "https://bricklink.com").searchParams.get("catString") || "";
		let ids = id.split(".");
		id = ids[ids.length - 1];

		return id;
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
		 * @property {(TypeList | undefined)[]} typeList - An array of type lists containing items.
		 *
		 * @typedef {Object} SearchResponse
		 * @property {Result} result - The result object containing the data.
		 * @property {number} returnCode - Return code indicating success or failure.
		 * @property {string} returnMessage - Message describing the result.
		 */
		/** @type {SearchResponse} */
		let { result, returnCode, returnMessage } = await response.json();
		if (returnCode != 0) throw new Error(returnMessage);

		/** @type {Item[]} */
		let items = [];
		if (result.typeList[0]) items = result.typeList[0].items;

		for (let item of items) {
			let themeIds = item.strCategory.split(".");
			let themeId = themeIds[themeIds.length - 1];

			legoSets.push({
				setNumber: item.strItemNo,
				name: item.strItemName,
				themeId
			});
		}

		return legoSets;
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
			throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
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
		await this.downloadImage(`https://img.bricklink.com/ItemImage/SN/0/${setNumber}.png`, path.join(SET_IMAGES_PATH, `${setNumber}.png`));
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
		await this.downloadImage(`https://img.bricklink.com/ItemImage/PN/${colorId}/${pieceId}.png`, path.join(PIECE_IMAGES_PATH, colorId.toString(), `${pieceId}.png`));
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
		await this.downloadImage(`https://img.bricklink.com/ItemImage/MN/0/${minifigId}.png`, path.join(MINIFIG_IMAGES_PATH, `${minifigId}.png`));
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
				this.downloadLegoPieceImage(setPiece.brickLinkId, setPiece.color?.brickLinkId || 0)
			),
			legoSet.normalPieces.forEach((setPiece) =>
				this.downloadLegoPieceImage(setPiece.brickLinkId, setPiece.color?.brickLinkId || 0)
			),
			legoSet.minifigPieces.forEach((setPiece) =>
				this.downloadLegoPieceImage(setPiece.brickLinkId, setPiece.color?.brickLinkId || 0)
			),
			legoSet.counterpartPieces.forEach((setPiece) =>
				this.downloadLegoPieceImage(setPiece.brickLinkId, setPiece.color?.brickLinkId || 0)
			),
			legoSet.minifigs.forEach((setPiece) => this.downloadMinifigImage(setPiece.brickLinkId))
		]);
	}
}
