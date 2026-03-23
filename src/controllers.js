// Imports
import { ClassMap } from "./classes.js";
import database from "./database/database.js";
import { webScrapper } from "./main.js";
import { LegoColor, LegoPiece, LegoSet, LegoSetPiece, LegoSetTheme } from "./models.js";


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
			this.add(
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

		return this.set(
			databaseId,
			new LegoColor(databaseId, bricklinkId, bricklinkName, legoId, legoName)
		);
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

		return this.set(
			databaseId,
			new LegoSetTheme(databaseId, bricklinkId, bricklinkName)
		);
	}

	/**
	 * @param {string} bricklinkId
	 */
	getByBricklinkId(bricklinkId) {
		for (let legoSetTheme of this.values()) {
			if (legoSetTheme.bricklinkId = bricklinkId) return legoSetTheme;
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
			this.set(legoPiece.databaseId, new LegoPiece(
				legoPiece.databaseId,
				legoPiece.bricklinkId,
				legoPiece.bricklinkName,
				/** @type {LegoColor} */(colors.get(legoPiece.colorId)),
				legoPiece.bricklinkCategory
			));
		}

		if (this.size > 0) return this;

		return this;
	}

	/**
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 * @param {LegoColor} color
	 * @param {string} bricklinkCategory
	 */
	async add(bricklinkId, bricklinkName, color, bricklinkCategory) {
		let databaseId = await database.addLegoPiece(bricklinkId, bricklinkName, color, bricklinkCategory);

		return this.set(databaseId, new LegoPiece(
			databaseId,
			bricklinkId,
			bricklinkName,
			color,
			bricklinkCategory
		));
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
	/**
	 * @param {Iterable<readonly [number, LegoSetPiece]>} [iterable]
	 */
	constructor(iterable) {
		super(LegoSetPiece, iterable);
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
	 * @param {string} setNumber
	 */
	getBySetNumber(setNumber) {
		for (let legoSet of this.values()) {
			if (legoSet.setNumber === setNumber) return legoSet;
		}

		return null;
	}
}
