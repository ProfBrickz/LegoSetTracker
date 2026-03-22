// Imports
import { ClassMap } from "./classes.js";
import database from "./database/database.js";
import { webScrapper } from "./main.js";
import { LegoColor, LegoPiece, LegoSet, LegoSetPiece } from "./models.js";


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


/** @extends {Map<string, string>} */
export class LegoSetThemes extends Map {
	async init() {
		let databaseThemes = await database.getLegoSetThemes();

		for (let { bricklinkId, bricklinkName } of databaseThemes) {
			this.set(bricklinkId, bricklinkName);
		}

		if (this.size > 0) return this;

		let webThemes = await webScrapper.getLegoSetThemes();

		for (let { bricklinkId, bricklinkName } of webThemes) {
			this.add(bricklinkId, bricklinkName);
		}

		return this;
	}

	/**
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 */
	async add(bricklinkId, bricklinkName) {
		await database.addLegoSetTheme(bricklinkId, bricklinkName);

		return this.set(bricklinkId, bricklinkName);
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
