// Imports
import axios from 'axios'
import * as inquirer from '@inquirer/prompts'
import fs from 'fs'
import { JSDOM } from "jsdom"
import * as childProcess from "child_process";

// Settings
const settings = JSON.parse(fs.readFileSync('settings.json'))

// Variables
let set = {
	name: null,
	id: null,
	parts: []
}
let sections = {}
let document = null
let tbody

// Classes
class Part {
	constructor(brickLinkId, name, imgUrl, amountNeeded, amountFound = 0) {
		this.brickLinkId = brickLinkId
		this.name = name
		this.imgUrl = imgUrl || ''

		this.imgPath = ''
		if (this.imgUrl) {
			let fileName = name.replace(/[\(\)\\/]/g, '').replace('  ', ' ')

			this.imgPath = `images/${fileName}.${imgUrl.split('.').at(-1)}`
		}

		this.amountNeeded = amountNeeded
		this.amountFound = amountFound
	}
}

// Functions
function getSection(section) {
	let categoryStartIndex = null
	let categoryEndIndex = null

	for (let category of catagories) {
		let index = rows.indexOf(category)

		if (category.textContent == section) {
			categoryStartIndex = index + 1
			continue
		}

		if (categoryStartIndex) {
			if (category.textContent == 'Parts:') {
				categoryStartIndex = index + 1
				continue
			}

			categoryEndIndex = index

			break
		}
	}

	if (categoryStartIndex && !categoryEndIndex) categoryEndIndex = rows.length - 1

	let categoryRows = rows.slice(categoryStartIndex, categoryEndIndex)

	return getParts(categoryRows)
}

function getParts(rows) {
	let parts = []

	for (let row of rows) {
		let brickLinkId = row.querySelector('td:nth-of-type(3) a').textContent.trim()
		let imageUrl = row.querySelector('td:nth-of-type(1) img').src
		let name = row.querySelector('td:nth-of-type(4) b').textContent.trim()
		// Remove repeated spaces
		name = name.replace(/\s+/g, ' ')

		let amountNeeded = Number.parseInt(row.querySelector('td:nth-of-type(2)').textContent)

		parts.push(new Part(brickLinkId, name, imageUrl, amountNeeded, 0))
	}

	return parts
}

// Main
if (!fs.existsSync('images')) fs.mkdirSync('images')
if (!fs.existsSync('sets')) fs.mkdirSync('sets')

try {
	let setId = await inquirer.input({
		message: 'What lego set do you want to find? (the set number)',
		default: settings.lastSet,
		validate: async (input) => {
			try {
				if (!input) return 'Enter a set number'
				if (Number.isInteger(Number(input))) input += '-1'

				const response = await axios.get(`https://www.bricklink.com/catalogItemInv.asp?S=${input}&viewType=P`, {
					headers: {
						'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
					}
				})

				document = (new JSDOM(response.data, {
					contentType: "text/html",
					'userAgent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
				})).window.document

				let setNameElement = document.querySelector('tbody center font b')
				tbody = document.querySelector('form > table tbody')

				if (!setNameElement) return 'Set not found.'
				if (!tbody) return 'Set does not have inventory yet.'

				set.id = input
				set.name = setNameElement.textContent.trim()

				return true
			} catch (error) {
				return error.message
			}
		}
	})

	if (Number.isInteger(Number(setId))) setId += '-1'

	settings.lastSet = setId


	let changeSettings = await inquirer.select({
		message: 'Do you want to go to change settings?',
		choices: [
			{ name: 'Yes', value: true },
			{ name: 'No', value: false }
		],
		default: false
	})

	if (changeSettings) {
		for (let key of Object.keys(settings.include)) {
			settings.include[key] = await inquirer.select({
				message: `Do you want to include ${key}?`,
				choices: [
					{ name: 'Yes', value: true },
					{ name: 'No', value: false }
				],
				default: settings.include[key]
			})
		}
	}
} catch (error) {
	if (error instanceof inquirer.ExitPromptError) {

		console.error('Prompt was closed unexpectedly.');
	} else {
		console.error('An unexpected error occurred:', error);
	}

	// stop program
	process.exit(0)
}


fs.writeFileSync('settings.json', JSON.stringify(settings))


let rows = Array.prototype.slice.call(tbody.children)
let catagories = document.querySelectorAll('form > table tr[BGCOLOR="#000000"], form > table tr[BGCOLOR="#C0C0C0"]')

sections.regularItems = getSection('Regular Items:')
sections.minifigures = getSection('Minifigures:')
sections.extraItems = getSection('Extra Items:')
sections.counterParts = getSection('Counterparts:')
sections.alternateItems = getSection('Alternate Items:')

let normalParts = structuredClone(sections.regularItems)

if (!settings.include.stickerSheet) {
	normalParts = normalParts.filter((part) => !part.brickLinkId.includes('stk'))
}

set.parts = set.parts.concat(normalParts)

if (settings.include.stickerParts) {
	let stickerParts = structuredClone(sections.counterParts)

	stickerParts = stickerParts.filter((part) => part.brickLinkId.includes('pb'))

	set.parts = set.parts.concat(stickerParts)

	for (let stickerPart of stickerParts) {
		let baseId = stickerPart.brickLinkId.split('pb')[0]
		let regularPart = set.parts.find(part =>
			part.brickLinkId === baseId ||
			part.brickLinkId === baseId + 'b'
		)
		if (!regularPart) continue

		if (regularPart) regularPart.amountNeeded -= stickerPart.amountNeeded

		if (regularPart.amountNeeded === 0) {
			set.parts = set.parts.filter(part => part.brickLinkId !== baseId)
		}
	}
}

if (settings.include.minifigures) {
	let minifigures = structuredClone(sections.minifigures)

	set.parts = set.parts.concat(minifigures)
}

for (let part of set.parts) {
	if (!part.imgPath) continue

	while (!fs.existsSync(part.imgPath)) {
		console.log(`downloading ${part.Name} Image`);

		let response = await axios.get(
			part.imgUrl,
			{
				headers: {
					'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
				},
				responseType: 'stream'
			}
		)

		response.data.pipe(
			fs.createWriteStream(
				part.imgPath
			)
		)
	}
}

fs.writeFileSync('set.json', JSON.stringify(set))

childProcess.execSync('dotnet run excelMaker/app.cs --project ./excelMaker/')