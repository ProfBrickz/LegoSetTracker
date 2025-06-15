// Imports
import * as inquirer from '@inquirer/prompts';
import axios from 'axios';
import * as childProcess from 'child_process';
import fs from 'fs';
import { JSDOM } from 'jsdom';
import { finished } from 'stream';
import { promisify } from 'util';

const finishedPromise = promisify(finished);

/**
 * Application settings loaded from settings.json
 * @type {Object}
 */
const settings = JSON.parse(fs.readFileSync('settings.json', 'utf-8'));

/**
 * Current set being processed
 * @type {Object}
 * @property {string|null} name - The name of the LEGO set
 * @property {number|null} amount - Quantity of the set being collected
 * @property {string|null} id - The BrickLink ID of the set
 * @property {Array<Part>} parts - List of parts in the set
 */
let set = {
	name: null,
	amount: null,
	id: null,
	parts: []
};

/**
 * Sections of parts categorized by type
 * @type {Object.<string, Array<Part>>}
 */
let sections = {};

/**
 * DOM document from BrickLink page
 * @type {Document}
 */
let document;

/**
 * Table body element containing parts information
 * @type {HTMLTableSectionElement}
 */
let tbody;

/**
 * Represents a LEGO part with its properties
 * @class
 */
class Part {
	/**
	 * Creates a new Part instance
	 *
	 * @param {string} brickLinkId - The BrickLink ID of the part
	 * @param {string} name - The name of the part
	 * @param {string} imgUrl - URL to the part image
	 * @param {number} amountNeeded - Number of this part needed
	 * @param {number} [amountFound=0] - Number of this part already found
	 */
	constructor(brickLinkId, name, imgUrl, amountNeeded, amountFound = 0) {
		this.brickLinkId = brickLinkId;
		this.name = name;
		this.imgUrl = imgUrl || '';

		this.imgPath = '';
		if (this.imgUrl) {
			let fileName = name.replace(/[\(\)\\/]/g, '').replace('  ', ' ');

			this.imgPath = `images/${fileName}.${imgUrl.split('.').at(-1)}`;
		}

		this.amountNeeded = amountNeeded;
		this.amountFound = amountFound;
	}
}

/**
 * Extracts parts from a specific section of the BrickLink page
 *
 * @param {string} section - The section name to extract parts from
 * @returns {Array<Part>} Array of parts from the specified section
 */
function getSection(section) {
	let categoryStartIndex = null;
	let categoryEndIndex = null;

	for (let category of categories) {
		let index = rows.indexOf(category);

		if (category.textContent == section) {
			categoryStartIndex = index + 1;
			continue;
		}

		if (categoryStartIndex) {
			if (category.textContent == 'Parts:') {
				categoryStartIndex = index + 1;
				continue;
			}

			categoryEndIndex = index;

			break;
		}
	}

	if (categoryStartIndex && !categoryEndIndex) categoryEndIndex = rows.length - 1;

	let categoryRows = rows.slice(categoryStartIndex, categoryEndIndex);

	return getParts(categoryRows);
}

/**
 * Extracts part information from table rows
 *
 * @param {Array<HTMLElement>} rows - The table rows containing part data
 * @returns {Array<Part>} Array of parsed parts
 */
function getParts(rows) {
	let parts = [];

	for (let row of rows) {
		/** @type {HTMLAnchorElement | null} */
		let brickLinkIdElement = row.querySelector('td:nth-of-type(3) a');
		/** @type {HTMLImageElement | null} */
		let imageElement = row.querySelector('td:nth-of-type(1) img');
		/** @type {HTMLElement | null} */
		let nameElement = row.querySelector('td:nth-of-type(4) b');
		/** @type {HTMLTableCellElement | null} */
		let amountNeededElement = row.querySelector('td:nth-of-type(2)');

		// Skip this row if any required element is missing
		if (
			!brickLinkIdElement
			|| !imageElement
			|| !nameElement
			|| !amountNeededElement
			|| !amountNeededElement.textContent
			|| !nameElement.textContent
		) {
			continue;
		}

		let brickLinkId = amountNeededElement.textContent.trim();
		let imageUrl = imageElement.src;
		let name = nameElement.textContent.trim();
		// Remove repeated spaces
		name = name.replace(/\s+/g, ' ');

		let amountNeeded = Number.parseInt(amountNeededElement.textContent);

		parts.push(new Part(brickLinkId, name, imageUrl, amountNeeded * set.amount, 0));
	}

	return parts;
}

// Main
if (!fs.existsSync('images')) fs.mkdirSync('images');
if (!fs.existsSync('sets')) fs.mkdirSync('sets');

try {
	/**
	 * Prompt user for set ID and validate it exists on BrickLink
	 */	// Variables to store outside the validation function
	let pageDocument;
	let pageTbody;

	let setId = await inquirer.input({
		message: 'What lego set do you want to find? (the set number)',
		default: settings.lastSet,
		validate: async (input) => {
			try {
				if (!input) return 'Enter a set number';
				if (Number.isInteger(Number(input))) input += '-1';

				const response = await axios.get(`https://www.bricklink.com/CatalogItemInv.asp?S=${input}&viewType=P`, {
					headers: {
						'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
					}
				});

				// Store document in a temporary variable within the validation scope
				const tempDocument = (new JSDOM(response.data, {
					contentType: "text/html",
					'userAgent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
				})).window.document;

				let setNameElement = tempDocument.querySelector('tbody center font b');
				if (!setNameElement || !setNameElement.textContent) {
					throw new Error('Set not found.');
				}

				const tempTbody = /** @type {HTMLTableSectionElement} */ (tempDocument.querySelector('form > table tbody'));

				if (!setNameElement) return 'Set not found.';
				if (!tempTbody) return 'Set does not have inventory yet.';
				// Save to outer scope for use after validation is complete
				pageDocument = tempDocument;
				pageTbody = tempTbody;

				set.id = input;
				set.name = setNameElement.textContent.trim();

				// Return true to indicate validation passed
				return true;
			} catch (error) {
				return error.message;
			}
		}
	});
	if (Number.isInteger(Number(setId))) setId += '-1';

	// Assign global variables after validation is complete
	if (!pageDocument || !pageTbody) {
		throw new Error('Failed to retrieve set information. Please try again.');
	}

	document = pageDocument;
	tbody = pageTbody;

	// Confirm the selected set with the user
	const confirmSet = await inquirer.confirm({
		message: `Is this the set you want to find? ${set.name} (${set.id})`,
		default: true
	});

	// If user doesn't confirm, exit the process
	if (!confirmSet) {
		console.log("Set selection cancelled. Exiting...");
		process.exit(0);
	}

	settings.lastSet = setId;

	/**
	 * Prompt user for set quantity
	 */
	let setAmount = await inquirer.input({
		message: 'How many of this set do you want to find?',
		default: '1',
		validate: (input) => {
			if (Number.isInteger(Number(input))) return true;
			return 'Enter an integer';
		}
	});

	set.amount = Number(setAmount);

	/**
	 * Prompt user to change settings
	 */
	let changeSettings = await inquirer.select({
		message: 'Do you want to go to change settings?',
		choices: [
			{ name: 'Yes', value: true },
			{ name: 'No', value: false }
		],
		default: false
	});

	if (changeSettings) {
		for (let key of Object.keys(settings.include)) {
			settings.include[key] = await inquirer.select({
				message: `Do you want to include ${key}?`,
				choices: [
					{ name: 'Yes', value: true },
					{ name: 'No', value: false }
				],
				default: settings.include[key]
			});
		}
	}
} catch (error) {
	// Log detailed error info
	console.log(error);

	if (error.message === 'Prompt was closed') {
		console.error('Prompt was closed unexpectedly.');
	} else {
		console.error('An unexpected error occurred:', error.message);
	}

	// stop program
	process.exit(0);
}

// Save updated settings
fs.writeFileSync('settings.json', JSON.stringify(settings));

// Extract part information from the HTML
let rows = Array.prototype.slice.call(tbody.children);
let categories = document.querySelectorAll('form > table tr[BGCOLOR="#000000"], form > table tr[BGCOLOR="#C0C0C0"]');

sections.regularItems = getSection('Regular Items:');
sections.minifigures = getSection('Minifigures:');
sections.extraItems = getSection('Extra Items:');
sections.counterParts = getSection('Counterparts:');
sections.alternateItems = getSection('Alternate Items:');

/**
 * Process regular parts based on settings
 */
let normalParts = structuredClone(sections.regularItems);

if (!settings.include.stickerSheet) {
	normalParts = normalParts.filter((part) => !part.brickLinkId.includes('stk'));
}

set.parts = set.parts.concat(normalParts);

/**
 * Process sticker parts if enabled in settings
 */
if (settings.include.stickerParts) {
	let stickerParts = structuredClone(sections.counterParts);

	stickerParts = stickerParts.filter((part) => part.brickLinkId.includes('pb'));

	set.parts = set.parts.concat(stickerParts);

	for (let stickerPart of stickerParts) {
		let baseId = stickerPart.brickLinkId.split('pb')[0];
		let regularPart = set.parts.find(part =>
			part.brickLinkId === baseId ||
			part.brickLinkId === baseId + 'b'
		);
		if (!regularPart) continue;

		if (regularPart) regularPart.amountNeeded -= stickerPart.amountNeeded;

		if (regularPart.amountNeeded === 0) {
			set.parts = set.parts.filter(part => part.brickLinkId !== baseId);
		}
	}
}

/**
 * Add minifigures if enabled in settings
 */
if (settings.include.minifigures) {
	let minifigures = structuredClone(sections.minifigures);

	set.parts = set.parts.concat(minifigures);
}

/**
 * Download images for all parts
 */
for (let part of set.parts) {
	if (!part.imgPath) continue;

	while (!fs.existsSync(part.imgPath)) {
		console.log(`downloading ${part.name} Image`);

		let response = await axios.get(
			part.imgUrl,
			{
				headers: {
					'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
				},
				responseType: 'arraybuffer'
			}
		);

		fs.writeFileSync(part.imgPath, Buffer.from(response.data));
	}
}

// Save set information to JSON file
fs.writeFileSync('set.json', JSON.stringify(set));

// Generate Excel file using ExcelMaker tool
console.log('Creating Excel File...');
childProcess.execSync('dotnet run excelMaker/app.cs --project ./excelMaker/');
