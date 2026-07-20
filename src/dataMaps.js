// Imports
import { ClassMap } from "./classes.js";
import database from "./database/database.js";
import { LegoColor, LegoPiece, LegoSet, LegoSetPiece, LegoSetTheme, StickeredLegoPiece, StickeredLegoSetPiece } from "./models.js";
import WebScrapper from "./webScrapper.js";
/** @import { LegoSetPieceType } from "./types.js" */


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
	 * @param {WebScrapper} webScrapper
	 */
	async init(webScrapper) {
		let databaseColors = await database.getLegoColors();

		for (let color of databaseColors) {
			this.set(color.databaseId, new LegoColor(
				color.databaseId,
				color.brickLinkId,
				color.brickLinkName,
				color.legoId,
				color.legoName
			));
		}

		if (this.size > 0) return this;

		let webColors = await webScrapper.getColors();

		for (let color of webColors) {
			await this.add(
				color.brickLinkId,
				color.brickLinkName,
				color.legoId,
				color.legoName
			);
		}

		return this;
	}

	/**
	 * @param {number} brickLinkId
	 * @param {string} brickLinkName
	 * @param {number | null} legoId
	 * @param {string | null} legoName
	 */
	async add(brickLinkId, brickLinkName, legoId, legoName) {
		let databaseId = await database.addLegoColor(brickLinkId, brickLinkName, legoId, legoName);

		let legoColor = new LegoColor(databaseId, brickLinkId, brickLinkName, legoId, legoName);

		this.set(databaseId, legoColor);
		return legoColor;
	}

	/**
	 * @param {number} brickLinkId
	 */
	getByBricklinkId(brickLinkId) {
		for (let legoColor of this.values()) {
			if (legoColor.brickLinkId == brickLinkId) return legoColor;
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

	/**
	 * Initializes the themes, by first checking the database then BrickLink
	 * @param {WebScrapper} webScrapper
	 */
	async init(webScrapper) {
		let databaseThemes = await database.getLegoSetThemes();

		for (let { databaseId, brickLinkId, brickLinkName } of databaseThemes) {
			this.set(databaseId, new LegoSetTheme(databaseId, brickLinkId, brickLinkName));
		}

		if (this.size > 0) return this;

		let webThemes = await webScrapper.getLegoSetThemes();

		for (let { brickLinkId, brickLinkName } of webThemes) {
			await this.add(brickLinkId, brickLinkName);
		}

		return this;
	}

	/**
	 * @param {string} brickLinkId
	 * @param {string} brickLinkName
	 */
	async add(brickLinkId, brickLinkName) {
		let databaseId = await database.addLegoSetTheme(brickLinkId, brickLinkName);

		let legoSetTheme = new LegoSetTheme(databaseId, brickLinkId, brickLinkName);

		this.set(databaseId, legoSetTheme);
		return legoSetTheme;
	}

	/**
	 * @param {string} brickLinkId
	 */
	getByBricklinkId(brickLinkId) {
		for (let legoSetTheme of this.values()) {
			if (legoSetTheme.brickLinkId == brickLinkId) return legoSetTheme;
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
		// let { pieces, stickeredRelations } = await database.getLegoPieces();
		let pieces = await database.getLegoPieces();
		let stickeredRelations = await database.getStickerRelations();

		// Initialize all base pieces
		for (let legoPiece of pieces) {
			let color = null;
			if (legoPiece.colorId != null) color = colors.get(legoPiece.colorId);
			if (typeof color == "undefined") throw new Error(`There is no Lego color with id ${legoPiece.colorId} in database.`);

			this.set(legoPiece.databaseId, new LegoPiece(
				legoPiece.databaseId,
				legoPiece.brickLinkId,
				legoPiece.brickLinkName,
				color,
				legoPiece.brickLinkCategory
			));
		}

		// Initialize stickered pieces by looking up their base pieces
		for (let relation of stickeredRelations) {
			let baseLegoPiece = this.get(relation.baseLegoPieceId);
			if (!baseLegoPiece) continue;

			let stickeredLegoPiece = this.get(relation.stickeredLegoPieceId);
			if (!stickeredLegoPiece) continue;

			this.set(stickeredLegoPiece.databaseId, new StickeredLegoPiece(
				stickeredLegoPiece.databaseId,
				stickeredLegoPiece.brickLinkId,
				stickeredLegoPiece.brickLinkName,
				stickeredLegoPiece.color,
				stickeredLegoPiece.brickLinkCategory,
				baseLegoPiece
			));
		}

		return this;
	}

	/**
	 * @param {string} brickLinkId
	 * @param {string} brickLinkName
	 * @param {LegoColor | null} color
	 * @param {string} brickLinkCategory
	 */
	async add(brickLinkId, brickLinkName, color, brickLinkCategory) {
		let databaseId = await database.addLegoPiece(brickLinkId, brickLinkName, color, brickLinkCategory);

		let legoPiece = new LegoPiece(
			databaseId,
			brickLinkId,
			brickLinkName,
			color,
			brickLinkCategory
		);

		this.set(databaseId, legoPiece);

		if (brickLinkName.includes("(Sticker)")) {
			let baseLegoPieceId = brickLinkId.split("pb")[0];
			let baseLegoPiece = this.getByBrickLinkIdAndColor(baseLegoPieceId, color);
			if (baseLegoPiece == null) return null;

			// Convert to StickeredLegoPiece and update in map
			let stickeredPiece = new StickeredLegoPiece(
				databaseId,
				brickLinkId,
				brickLinkName,
				color,
				brickLinkCategory,
				baseLegoPiece
			);
			this.set(databaseId, stickeredPiece);

			// Record the relationship in database
			await database.addStickeredLegoPiece(stickeredPiece);
		}

		return legoPiece;
	}

	/**
	 * @param {string} brickLinkId
	 * @param {LegoColor | null} color
	 */
	getByBrickLinkIdAndColor(brickLinkId, color) {
		for (let value of this.values()) {
			if (
				value.brickLinkId == brickLinkId
				&& value.color == color
			) return value;
		}

		return null;
	}

	/**
	 * @param {string} brickLinkId
	 * @param {string} brickLinkName
	 * @param {LegoColor | null} color
	 * @param {string} brickLinkCategory
	 */
	getDatabaseId(brickLinkId, brickLinkName, color, brickLinkCategory) {
		for (let value of this.values()) {
			if (
				value.brickLinkId == brickLinkId
				&& value.brickLinkName == brickLinkName
				&& value.color == color
				&& value.brickLinkCategory == brickLinkCategory
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

		if (legoPiece instanceof StickeredLegoPiece) {
			let baseLegoSetPiece = this.getBaseSetPiece(legoPiece);
			if (baseLegoSetPiece == null) return this;

			return this.set(databaseId, new StickeredLegoSetPiece(
				databaseId,
				legoPiece,
				baseLegoSetPiece,
				amountNeeded,
				amountFound
			));
		} else {
			return this.set(databaseId, new LegoSetPiece(
				databaseId,
				legoPiece,
				amountNeeded,
				amountFound
			));
		}
	}

	/**
	 * @param {StickeredLegoPiece} stickeredLegoPiece
	 */
	getBaseSetPiece(stickeredLegoPiece) {
		for (let legoSetPiece of this.values()) {
			if (
				legoSetPiece.brickLinkId == stickeredLegoPiece.brickLinkId
				&& legoSetPiece.color == stickeredLegoPiece.color
			) return legoSetPiece;
		}

		return null;
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

				let legoSetPieces;
				if (legoSetPiece.legoSetPieceType == "minifig") legoSetPieces = legoSet.minifigs;
				else if (legoSetPiece.legoSetPieceType == "extra") legoSetPieces = legoSet.extraPieces;
				else if (legoSetPiece.legoSetPieceType == "counterpart") legoSetPieces = legoSet.counterpartPieces;
				else legoSetPieces = legoSet.normalPieces;

				if (legoPiece instanceof StickeredLegoPiece) {
					let baseLegoSetPiece = legoSetPieces.getBaseSetPiece(legoPiece);
					if (baseLegoSetPiece == null) continue;

					legoSetPieces.set(legoSetPiece.databaseId, new StickeredLegoSetPiece(
						legoSetPiece.databaseId,
						legoPiece,
						baseLegoSetPiece,
						legoSetPiece.amountNeeded,
						legoSetPiece.amountFound
					));
				} else {
					legoSetPieces.set(legoSetPiece.databaseId, new LegoSetPiece(
						legoSetPiece.databaseId,
						legoPiece,
						legoSetPiece.amountNeeded,
						legoSetPiece.amountFound
					));
				}
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
