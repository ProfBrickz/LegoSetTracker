// Imports
import * as fs from 'fs';
import { JSDOM } from 'jsdom';
import "./types";



// Functions
/**
 * Fetches and parses the HTML content of a website.
 * 
 * @param {string} url
 * @returns {Promise<Document>}
 * @throws {Error} If the fetch operation fails or the response is not successful
 */
async function getWebpage(url) {
	let response = await fetch(url);

	if (!response.ok) {
		throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
	}

	let html = await response.text();
	let { document } = new JSDOM(html).window;
	return document;
}

/**
 * @param {string} setNumberInput The input value for the set number.
 * @returns {Promise<SetInfo>} A promise that resolves to a SetInfo object containing details about the
 */
async function getSetInfo(setNumberInput) {
	let document = await getWebpage(`https://www.bricklink.com/v2/catalog/catalogitem.page?S=${setNumberInput}`);

	// Extracting the theme from the document
	let themeElement = /** @type {HTMLElement|null} */ (document.querySelector("#content .innercontent table:first-of-type tr td:nth-child(1) :nth-child(3)"));

	if (!themeElement) {
		throw new Error(`Could not find the set ${setNumberInput}`);
	}

	let theme = themeElement.textContent.trim();

	// Extracting the name from the document
	let nameElement = /** @type {HTMLElement} */ (document.querySelector("#id_divBlock_Main table:first-of-type tr:first-of-type td:first-of-type h1"));
	let name = nameElement.textContent.trim();

	// Extracting the set number from the document
	let setNumberElement = /** @type {HTMLElement} */ (document.querySelector("#id_divBlock_Main table:first-of-type tr:first-of-type td:first-of-type span span"));
	let setNumber = setNumberElement.textContent.trim();

	// Extracting the year from the document
	let yearElement = /** @type {HTMLElement} */ (document.querySelector("#id_divBlock_Main table:first-of-type tr:nth-of-type(2) td:nth-of-type(2) table tr:first-of-type td:first-of-type a:first-of-type"));
	let year = Number.parseInt(yearElement.textContent.trim());

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

	return { name, setNumber, theme, year, pieceCount, minifigCount };
}

/**
 * @param {string} setNumber The input value for the set number.
 * @returns {Promise<SetPieceInfo>}
 */
async function getSetPieces(setNumber) {
	let document = await getWebpage(`https://www.bricklink.com/catalogItemInv.asp?S=${setNumber}&viewType=P&sortBy=0&bt=0&sortAsc=a`);
	let tbody = /** @type {HTMLTableSectionElement}*/ (document.querySelector("form > table tbody"));
	let rows = /** @type {HTMLTableRowElement[]} */ (Array.from(tbody.childNodes));
	let categories = /** @type {NodeListOf<HTMLTableRowElement>} */ (document.querySelectorAll('form > table tr[bgcolor="#000000"], form > table tr[bgcolor="#C0C0C0"]'));

	let normalPieces = getSection("Regular Items:", rows, categories);
	let minifigs = getSection("Minifigures:", rows, categories);
	let extraPieces = getSection("Extra Items:", rows, categories);
	let counterparts = getSection("Counterparts:", rows, categories);

	return { normalPieces, minifigs, extraPieces, counterparts };
}

/**
 * @param {string} section
 * @param {HTMLTableRowElement[]} rows
 * @param {NodeListOf<HTMLTableRowElement>} categories
 * @returns {SetPiece[]}
 */
function getSection(section, rows, categories) {
	/** @type {number|null} */
	let sectionStartIndex = null;
	/** @type {number|null} */
	let sectionEndIndex = null;

	for (let category of categories) {
		let index = rows.indexOf(category);

		if (category.textContent == section) {
			sectionStartIndex = index + 1;
			continue;
		}

		if (sectionStartIndex) {
			if (category.textContent == 'Parts:') {
				sectionStartIndex = index + 1;
				continue;
			}

			sectionEndIndex = index;
			break;
		}
	}

	if (sectionStartIndex && !sectionEndIndex) sectionEndIndex = rows.length;

	if (!sectionStartIndex || !sectionEndIndex) throw new Error(`Could not find ${section} in categories`);

	let categoryRows = rows.slice(sectionStartIndex, sectionEndIndex);

	return getSectionPieces(categoryRows);
}

/**
 * @param {HTMLTableRowElement[]} categoryRows The rows of the category to parse.
 * @returns {SetPiece[]}
 */
function getSectionPieces(categoryRows) {
	/** @type {SetPiece[]})[]} */
	let pieces = [];

	let i = 0;
	for (let row of categoryRows) {
		let bricklinkIdElement = /** @type {HTMLAnchorElement} */ (row.querySelector('td:nth-of-type(3) a'));
		let bricklinkId = bricklinkIdElement.textContent.trim();
		let imageUrlElement = /** @type {HTMLImageElement} */ (row.querySelector('td:nth-of-type(1) img'));
		let imageUrl = imageUrlElement.src;
		let nameElement = /** @type {HTMLElement} */ (row.querySelector('td:nth-of-type(4) b'));
		let nameAndColor = nameElement.textContent.trim();
		// Remove repeated spaces
		nameAndColor = nameAndColor.replace(/\s+/g, ' ');

		let categoryElement = /** @type {HTMLTableCellElement} */ (row.querySelector("td:nth-of-type(4) a:nth-of-type(3)"));
		let category = categoryElement.textContent.trim();

		let amountNeededElement = /** @type {HTMLTableCellElement} */ (row.querySelector('td:nth-of-type(2)'));
		let amountNeeded = Number.parseInt(amountNeededElement.textContent);

		let color = colors.find(
			(color) => nameAndColor
				.split(" ")
				.slice(0, 5)
				.join(" ")
				.includes(color.bricklinkName)
		) || null;
		let name = nameAndColor;
		if (color) name = nameAndColor.replace(color.bricklinkName, '').trim();

		pieces.push({
			piece: {
				bricklinkId,
				name,
				color,
				category,
				parentId: null,
				children: []
			},
			type: "counterpart",
			amountNeeded
		});
	}

	return pieces;
}

/**
 * @returns {Promise<Color[]>} A set of Colors
 */
export async function getColors() {
	let document = await getWebpage("https://v2.bricklink.com/en-us/catalog/color-guide");

	let sections = /** @type {NodeListOf<HTMLTableSectionElement>} */ (document.querySelectorAll(".color-list-wide-viewport_hideMobileViewport__5OSVt tbody"));
	/** @type {Color[]} */
	let colors = [];

	for (let section of sections) {
		for (let tr of section.children) {
			let bricklinkNameElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(2) p"));
			let bricklinkName = bricklinkNameElement.textContent.trim();

			let legoNameElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(2) span"));
			let [legoName, legoIdStr] = legoNameElement.textContent.trim().slice(12).split(" - ");
			let legoId = Number.parseInt(legoIdStr);

			let bricklinkIdElement = /** @type {HTMLParagraphElement} */(tr.querySelector("td:nth-of-type(8)"));
			let bricklinkId = Number.parseInt(bricklinkIdElement.textContent);

			colors.push({ bricklinkId, bricklinkName, legoId, legoName });
		}
	}

	return colors;
}

/**
 * @param {string} minifigId The ID of the minifig to get pieces for.
 * @returns {Promise<SetPiece[]>} 
 */
export async function getMinifigPieces(minifigId) {
	let document = await getWebpage(`https://www.bricklink.com/catalogItemInv.asp?M=${minifigId}&viewType=P&bt=0&sortBy=0&sortAsc=a`);
	let tbody = /** @type {HTMLTableSectionElement}*/ (document.querySelector("form > table tbody"));
	let rows = /** @type {HTMLTableRowElement[]} */ (Array.from(tbody.childNodes));
	let categories = /** @type {NodeListOf<HTMLTableRowElement>} */ (document.querySelectorAll('form > table tr[bgcolor="#000000"], form > table tr[bgcolor="#C0C0C0"]'));

	return getSection("Regular Items:", rows, categories);
}

/**
 * @param {string} pieceId The ID of the piece to get composite pieces for.
 * @param {Color} color The color of the piece to get composite pieces for.
 * @returns {Promise<SetPiece[]>} 
 */
export async function getCompositePiece(pieceId, color) {
	let document = await getWebpage(`https://www.bricklink.com/catalogItemInv.asp?P=${pieceId}&C=${color.bricklinkId}&viewType=P&bt=0&sortBy=0&sortAsc=a`);
	let tbody = /** @type {HTMLTableSectionElement}*/ (document.querySelector("form > table tbody"));
	let rows = /** @type {HTMLTableRowElement[]} */ (Array.from(tbody.childNodes));
	let categories = /** @type {NodeListOf<HTMLTableRowElement>} */ (document.querySelectorAll('form > table tr[bgcolor="#000000"], form > table tr[bgcolor="#C0C0C0"]'));

	let pieces = getSection("Regular Items:", rows, categories);

	for (let piece of pieces) {
		if (!piece.piece.color) {
			piece.piece.color = color;
		}
	}

	return pieces;
}

/**
 * @param {string} url
 * @param {string} downloadPath
 * @returns {Promise<void>}
 */
async function downloadImage(url, downloadPath) {
	let response = await fetch(url);

	if (!response.ok) {
		throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
	}

	let arrayBuffer = await response.arrayBuffer();

	fs.writeFile(downloadPath, Buffer.from(arrayBuffer), {}, (error) => {
		if (error) {
			console.error("Error writing file: ", error);
			return;
		}
	});
}

/**
 * @param {string} setNumber
 * @returns {Promise<void>}
*/
async function downloadSetImage(setNumber) {
	downloadImage(`https://img.bricklink.com/S/${setNumber}.jpg`, `./images/sets/${setNumber}.jpg`);
}

/**
 * @param {string} pieceId
 * @param {number} colorId
 * @returns {Promise<void>}
*/
async function downloadPieceImage(pieceId, colorId) {
	downloadImage(`https://img.bricklink.com/P/${colorId}/${pieceId}.jpg`, `./images/pieces/${colorId}/${pieceId}.jpg`);
}

/**
 * @param {string} minifigId 
 * @returns {Promise<void>}
 */
async function downloadMinifigImage(minifigId) {
	downloadImage(`https://img.bricklink.com/M/${minifigId}.jpg`, `./images/minifigs/${minifigId}.jpg`);
}

/**
 * @param {string} setNumber The input value for the set number.
 * @returns {Promise<LegoSet>} A promise that resolves to a LegoSet object containing details about the
 */
export async function getLegoSet(setNumber) {
	let setInfo = await getSetInfo(setNumber);
	let setPieces = await getSetPieces(setNumber);

	return {
		...setInfo,
		...setPieces
	};
}

/**
 * Downloads images for a list of pieces.
 * @param {LegoSet} legoSet
 * @returns {Promise<void>} - A promise that resolves when all images have been downloaded.
 * @throws {Error} - If any image fails to download.
 */
export async function downloadSetImages(legoSet) {
	Promise.all([
		downloadSetImage(legoSet.setNumber),
		legoSet.normalPieces.map((setPiece) =>
			downloadPieceImage(setPiece.piece.bricklinkId, setPiece.piece.color?.bricklinkId || 0)
		),
		legoSet.counterparts.map((setPiece) =>
			downloadPieceImage(setPiece.piece.bricklinkId, setPiece.piece.color?.bricklinkId || 0)
		),
		legoSet.minifigs.map((setPiece) => downloadMinifigImage(setPiece.piece.bricklinkId))
	]);
}


// temp
/** @type {Color[]} */
let colors = [];
