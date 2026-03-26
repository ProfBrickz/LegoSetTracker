// Imports
import { ClassMap } from "./classes.js";
import database from "./database/database.js";
import { webScrapper } from "./main.js";
import { LegoColor, LegoPiece, LegoSet, LegoSetPiece, LegoSetTheme } from "./models.js";
/** @import {LegoSetPieceType} from "./types.js" */


// Classes
/** @extends {ClassMap<number, LegoColor>}  */
export class LegoColors extends ClassMap {
	/**
	 * @param {Iterable<readonly [number, LegoColor]>} [iterable]
	 */
	constructor(iterable) {
		super(LegoColor, iterable);
	}

	/**
	 * Initializes the colors, by first checking the database then BrickLink
	 */
	async init() {
		let databaseColors = await database.getLegoColors();

		for (let color of databaseColors) {
			this.set(color.databaseId, new LegoColor(
				color.databaseId,
				color.bricklinkId,
				color.bricklinkName,
				color.legoId,
				color.legoName
			));
		}

		if (this.size > 0) return this;

		let webColors = await webScrapper.getColors();

		for (let color of webColors) {
			await this.add(
				color.bricklinkId,
				color.bricklinkName,
				color.legoId,
				color.legoName
			);
		}

		return this;
	}

	/**
	 * @param {number} bricklinkId
	 * @param {string} bricklinkName
	 * @param {number | null} legoId
	 * @param {string | null} legoName
	 */
	async add(bricklinkId, bricklinkName, legoId, legoName) {
		let databaseId = await database.addLegoColor(bricklinkId, bricklinkName, legoId, legoName);

		let legoColor = new LegoColor(databaseId, bricklinkId, bricklinkName, legoId, legoName);

		this.set(databaseId, legoColor);
		return legoColor;
	}

	/**
	 * @param {number} bricklinkId
	 */
	getByBricklinkId(bricklinkId) {
		for (let legoColor of this.values()) {
			if (legoColor.bricklinkId == bricklinkId) return legoColor;
		}

		return null;
	}
}


/** @extends {ClassMap<number, LegoSetTheme>} */
export class LegoSetThemes extends ClassMap {
	/**
	 * @param {Iterable<readonly [number, LegoSetTheme]>} [iterable]
	 */
	constructor(iterable) {
		super(LegoSetTheme, iterable);
	}

	async init() {
		let databaseThemes = await database.getLegoSetThemes();

		for (let { databaseId, bricklinkId, bricklinkName } of databaseThemes) {
			this.set(databaseId, new LegoSetTheme(databaseId, bricklinkId, bricklinkName));
		}

		if (this.size > 0) return this;

		let webThemes = await webScrapper.getLegoSetThemes();

		for (let { bricklinkId, bricklinkName } of webThemes) {
			await this.add(bricklinkId, bricklinkName);
		}

		return this;
	}

	/**
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 */
	async add(bricklinkId, bricklinkName) {
		let databaseId = await database.addLegoSetTheme(bricklinkId, bricklinkName);

		let legoSetTheme = new LegoSetTheme(databaseId, bricklinkId, bricklinkName);

		this.set(databaseId, legoSetTheme);
		return legoSetTheme;
	}

	/**
	 * @param {string} bricklinkId
	 */
	getByBricklinkId(bricklinkId) {
		for (let legoSetTheme of this.values()) {
			if (legoSetTheme.bricklinkId == bricklinkId) return legoSetTheme;
		}

		return null;
	}
}


/** @extends {ClassMap<number, LegoPiece>}  */
export class LegoPieces extends ClassMap {
	/**
	 * @param {Iterable<readonly [number, LegoPiece]>} [iterable]
	 */
	constructor(iterable) {
		super(LegoPiece, iterable);
	}

	/**
	 * @param {LegoColors} colors
	 */
	async init(colors) {
		let databaseLegoPieces = await database.getLegoPieces();

		for (let legoPiece of databaseLegoPieces) {
			let color = null;
			if (legoPiece.colorId) color = (colors.get(legoPiece.colorId));
			if (!color) throw new Error(`There is no Lego color with id ${legoPiece.colorId} in database.`);

			this.set(legoPiece.databaseId, new LegoPiece(
				legoPiece.databaseId,
				legoPiece.bricklinkId,
				legoPiece.bricklinkName,
				color,
				legoPiece.bricklinkCategory
			));
		}

		return this;
	}

	/**
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 * @param {LegoColor | null} color
	 * @param {string} bricklinkCategory
	 */
	async add(bricklinkId, bricklinkName, color, bricklinkCategory) {
		let databaseId = await database.addLegoPiece(bricklinkId, bricklinkName, color, bricklinkCategory);

		let legoPiece = new LegoPiece(
			databaseId,
			bricklinkId,
			bricklinkName,
			color,
			bricklinkCategory
		);

		this.set(databaseId, legoPiece);
		return legoPiece;
	}

	/**
	 *
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 * @param {LegoColor | null} color
	 * @param {string} bricklinkCategory
	 */
	getDatabaseId(bricklinkId, bricklinkName, color, bricklinkCategory) {
		for (let value of this.values()) {
			if (
				value.bricklinkId == bricklinkId
				&& value.bricklinkName == bricklinkName
				&& value.color == color
				&& value.bricklinkCategory == bricklinkCategory
			) return value.databaseId;
		}

		return null;
	}
}


/** @extends {ClassMap<number, LegoSetPiece>}  */
export class LegoSetPieces extends ClassMap {
	/** @type {LegoSetPieceType} */
	legoSetPieceType;

	/**
	 * @param {LegoSetPieceType} legoSetPieceType
	 * @param {Iterable<readonly [number, LegoSetPiece]>} [iterable]
	 */
	constructor(legoSetPieceType, iterable) {
		super(LegoSetPiece, iterable);

		this.legoSetPieceType = legoSetPieceType;
	}

	/**
	 * @param {LegoSet} legoSet
	 * @param {LegoPiece} legoPiece
	 * @param {number} amountNeeded
	 * @param {number} amountFound
	*/
	async add(legoSet, legoPiece, amountNeeded, amountFound) {
		let databaseId = await database.addLegoSetPiece(
			legoSet.databaseId,
			legoPiece.databaseId,
			this.legoSetPieceType,
			amountNeeded,
			amountFound
		);

		return this.set(databaseId, new LegoSetPiece(
			databaseId,
			legoPiece,
			amountNeeded,
			amountFound
		));
	}
}


/** @extends {ClassMap<number, LegoSet>}  */
export class LegoSets extends ClassMap {
	/**
	 * @param {Iterable<readonly [number, LegoSet]>} [iterable]
	 */
	constructor(iterable) {
		super(LegoSet, iterable);
	}

	/**
	 * @param {LegoSetThemes} legoSetThemes
	 * @param {LegoPieces} legoPieces
	 */
	async init(legoSetThemes, legoPieces) {
		let databaseLegoSets = await database.getLegoSets();

		for (let databaseLegoSet of databaseLegoSets) {
			let theme = legoSetThemes.get(databaseLegoSet.themeId);
			if (!theme) throw new Error(`There is no theme with id ${databaseLegoSet.themeId} in database.`);

			let legoSet = new LegoSet(
				databaseLegoSet.databaseId,
				databaseLegoSet.setNumber,
				databaseLegoSet.name,
				theme,
				databaseLegoSet.releaseYear,
				databaseLegoSet.pieceCount,
				databaseLegoSet.minifigCount,
				databaseLegoSet.legoSetCount
			);

			this.set(databaseLegoSet.databaseId, legoSet);

			let databaseLegoSetPieces = await database.getLegoSetPieces(databaseLegoSet.databaseId);

			for (let legoSetPiece of databaseLegoSetPieces) {
				let legoPiece = legoPieces.get(legoSetPiece.legoPieceId);
				if (!legoPiece) throw new Error(`There is no Lego piece with the id ${legoSetPiece} in database.`);

				legoSet.normalPieces.set(
					legoSetPiece.databaseId,
					new LegoSetPiece(
						legoSetPiece.databaseId,
						legoPiece,
						legoSetPiece.amountNeeded,
						legoSetPiece.amountFound
					)
				);
			}
		}

		return this;
	}

	/**
	 * @param {string} setNumber
	 * @param {string} name
	 * @param {LegoSetTheme} theme
	 * @param {number} releaseYear
	 * @param {number} pieceCount
	 * @param {number} minifigCount
	 * @param {number} legoSetCount
	 */
	async add(
		setNumber,
		name,
		theme,
		releaseYear,
		pieceCount,
		minifigCount,
		legoSetCount,
	) {
		let databaseId = await database.addLegoSet(
			setNumber,
			name,
			theme.databaseId,
			releaseYear,
			pieceCount,
			minifigCount,
			legoSetCount
		);

		let legoSet = new LegoSet(
			databaseId,
			setNumber,
			name,
			theme,
			releaseYear,
			pieceCount,
			minifigCount,
			legoSetCount
		);

		this.set(databaseId, legoSet);
		return legoSet;
	}

	/**
	 * @param {string} setNumber
	 */
	getBySetNumber(setNumber) {
		for (let legoSet of this.values()) {
			if (legoSet.setNumber === setNumber) return legoSet;
		}

		return null;
	}
}
