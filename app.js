import axios from 'axios'
import * as excel from 'excel4node'
import fs from 'fs'
import inquirer from 'inquirer'
// import data from './settings.json' assert { type: "json" }

let settings = JSON.parse(fs.readFileSync('settings.json'))
const apiKey = settings.apiKey
const limit = settings.limit

let spreadsheet = [[
	'Done',
	'Amount Found',
	'Amount Needed',
	'BrickLink ID',
	'Lego ID',
	'Image',
	'Name'
]]
let styles = {}
let name = ''
let set = ''

styles.normal = {
	font: {
		color: '#000000',
		size: 14,
	},
	alignment: {
		horizontal: 'center',
		vertical: 'center',
		shrinkToFit: true,
		wrapText: true
	}
}
styles.header = {
	...styles.normal,
	font: {
		...styles.normal.font,
		bold: true
	},
	fill: {
		type: 'pattern',
		patternType: 'solid',
		fgColor: '#BDD7EE'
	}
}
styles.done = {
	...styles.normal,
	font: {
		...styles.normal.font,
		bold: true
	},
}
styles.found = {
	...styles.normal,
	fill: {
		type: 'pattern',
		patternType: 'solid',
		fgColor: '#C6E0B4'
	}
}
styles.info = {
	...styles.normal,
	fill: {
		type: 'pattern',
		patternType: 'solid',
		fgColor: '#EBEBEB'
	}
}

do {
	const answers = await inquirer.prompt({
		name: 'set',
		type: 'input',
		message: 'What lego set do you want to find? (the set number)',
	})
	try {
		let response = await axios.get(`https://rebrickable.com/api/v3/lego/sets/${answers.set}/?key=${apiKey}`)
		if (response.data) {
			set = answers.set
			name = response.data.name
		} else console.log('That set does not exist')
	} catch (error) {
		console.log(error.message)
	}
} while (!set)

if (!fs.existsSync('images/')) {
	fs.mkdirSync('images/')
}

try {
	let response = await axios.get(`https://rebrickable.com/api/v3/lego/sets/${set}/parts?key=${apiKey}&page_size=${limit}`)
	let results = response.data.results
	for (let result of results) {
		if (!result.is_spare) {
			// let id = result.element_id
			let id = result.part.part_img_url.split('/').at(-1).split('.')[0]
			spreadsheet.push({
				brickLinkId: result.part.external_ids.BrickLink + ' ' + result.color.external_ids.BrickLink.ext_descrs[0][0],
				legoId: id,
				amount: result.quantity,
				imgUrl: result.part.part_img_url,
				imgPath: 'images/' + result.part.part_img_url.split('/').at(-1),
				name: result.color.external_ids.BrickLink.ext_descrs[0][0] + ' ' + result.part.name
			})
		}
	}
} catch (error) {
	console.log(error.message)
}

let workbook = new excel.Workbook({
	defaultFont: {
		size: 12,
		name: 'Arial',
		color: '000000',
	}
})
let worksheet = workbook.addWorksheet(name, {
	sheetFormat: {
		defaultColWidth: 16,
		defaultRowHeight: 48
	}
})
worksheet.addConditionalFormattingRule(`A2:A${spreadsheet.length}`, {
	type: 'expression',
	priority: 1,
	formula: '=AND($B2<$C2,$B2>0)',
	style: workbook.createStyle({
		fill: {
			type: 'pattern',
			patternType: 'solid',
			bgColor: '#FFFF00'
		}
	}),
})
worksheet.addConditionalFormattingRule(`A2:A${spreadsheet.length}`, {
	type: 'expression',
	priority: 1,
	formula: '=$B2>$C2',
	style: workbook.createStyle({
		fill: {
			type: 'pattern',
			patternType: 'solid',
			bgColor: '#ED7D31'
		}
	}),
})
worksheet.addConditionalFormattingRule(`A2:A${spreadsheet.length}`, {
	type: 'expression',
	priority: 1,
	formula: '=$B2=$C2',
	style: workbook.createStyle({
		fill: {
			type: 'pattern',
			patternType: 'solid',
			bgColor: '#00B050'
		}
	}),
})
worksheet.addConditionalFormattingRule(`A2:A${spreadsheet.length}`, {
	type: 'expression',
	priority: 1,
	formula: '=$B2<=0',
	style: workbook.createStyle({
		fill: {
			type: 'pattern',
			patternType: 'solid',
			bgColor: '#FF0000'
		}
	}),
})
worksheet.row(1).freeze()
worksheet.column(7).setWidth(48)

for (let columnIndex = 0; columnIndex < spreadsheet[0].length; columnIndex++) {
	let column = spreadsheet[0][columnIndex]
	worksheet
		.cell(1, columnIndex + 1).string(column)
		.style(styles.header)
}

for (let rowIndex = 2; rowIndex < spreadsheet.length + 1; rowIndex++) {
	let row = spreadsheet[rowIndex - 1]
	try {
		worksheet
			//↑
			.cell(rowIndex, 1).formula(`=IF(B${rowIndex}<C${rowIndex},IF(B${rowIndex}>0,"—","✖"),IF(B${rowIndex}>C${rowIndex},"^",IF(B${rowIndex}=B${rowIndex},"✔","✖")))`)
			.style(styles.done)
		worksheet
			.cell(rowIndex, 2).number(0)
			.style(styles.found)
		worksheet
			.cell(rowIndex, 3).number(row.amount)
			.style(styles.info)
		worksheet
			.cell(rowIndex, 4).string(row.brickLinkId)
			.style(styles.info)
		worksheet
			.cell(rowIndex, 5).string(row.legoId)
			.style(styles.info)
		console.log(`Downloading ${row.legoId}`)
		while (!fs.existsSync(row.imgPath)) {
			let response = await axios.get(
				row.imgUrl,
				{ responseType: "stream" }
			)
			response.data.pipe(
				fs.createWriteStream(
					row.imgPath
				)
			)
		}
		worksheet.addImage({
			path: row.imgPath,
			type: 'picture',
			position: {
				type: 'twoCellAnchor',
				from: {
					col: 6,
					colOff: "11mm",
					row: rowIndex,
					rowOff: "1mm"
				},
				to: {
					col: 6,
					colOff: "29mm",
					row: rowIndex,
					rowOff: "19mm"
				}
			}
		})
		worksheet
			.cell(rowIndex, 7).string(row.name)
			.style(styles.info)
	} catch (error) {
		console.log(error.message)
	}
}

workbook.write(name + '.xlsx', (error, stats) => {
	if (error) {
		console.error(error)
	}
})