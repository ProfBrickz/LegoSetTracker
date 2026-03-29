// Imports
import path from "path";
import { MINIFIG_IMAGES_PATH, PIECE_IMAGES_PATH, SET_IMAGES_PATH } from "./constants.js";
import { LegoPieces, LegoSetPieces } from "./controllers.js";
import database from "./database/database.js";
/** @import { LegoSetPieceInfo } from "./types.js" */


// Classes
export class LegoColor {
	/** @type {string} */
	#brand = "";
	/** @type {number} */
	databaseId;
	/** @type {number} */
	bricklinkId;
	/** @type {string} */
	bricklinkName;
	/** @type {number | null} */
	legoId;
	/** @type {string | null} */
	legoName;

	/**
	 * @param {number} databaseId
	 * @param {number} bricklinkId
	 * @param {string} bricklinkName
	 * @param {number | null} legoId
	 * @param {string | null} legoName
	 */
	constructor(databaseId, bricklinkId, bricklinkName, legoId, legoName) {
		this.databaseId = databaseId;
		this.bricklinkId = bricklinkId;
		this.bricklinkName = bricklinkName;
		this.legoId = legoId;
		this.legoName = legoName;
	}
}

export class LegoPiece {
	/** @type {string} */
	#brand = "";
	/** @type {number} */
	databaseId;
	/** @type {string} */
	bricklinkId;
	/** @type {string} */
	bricklinkName;
	/** @type {LegoColor | null} */
	color;
	/** @type {string} */
	bricklinkCategory;

	/**
	 * @param {number} databaseId
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 * @param {LegoColor | null} color
	 * @param {string} bricklinkCategory
	*/
	constructor(databaseId, bricklinkId, bricklinkName, color, bricklinkCategory) {
		this.databaseId = databaseId;
		this.bricklinkId = bricklinkId;
		this.bricklinkName = bricklinkName;
		this.color = color;
		this.bricklinkCategory = bricklinkCategory;
	}

	getImagePath() {
		let colorId = this.color?.bricklinkId || 0;

		return path.join(PIECE_IMAGES_PATH, colorId.toString(), `${this.bricklinkId}.png`);
	}

	getMinifigImagePath() {
		return path.join(MINIFIG_IMAGES_PATH, `${this.bricklinkId}.png`);
	}
}

export class LegoSetPiece {
	/** @type {string} */
	#brand = "";
	/** @type {number} */
	databaseId;
	/** @type {LegoPiece} */
	#legoPiece;
	/** @type {number} */
	amountNeeded;
	/** @type {number} */
	amountFound;

	// Make the constructor with the jsdoc string
	/**
	 * @param {number} databaseId
	 * @param {LegoPiece} legoPiece
	 * @param {number} amountNeeded
	 * @param {number} [amountFound=0]
	 */
	constructor(databaseId, legoPiece, amountNeeded, amountFound = 0) {
		this.databaseId = databaseId;
		this.#legoPiece = legoPiece;
		this.amountNeeded = amountNeeded;
		this.amountFound = amountFound;
	}

	get bricklinkId() {
		return this.#legoPiece.bricklinkId;
	}

	get bricklinkName() {
		return this.#legoPiece.bricklinkName;
	}

	get color() {
		return this.#legoPiece.color;
	}

	get bricklinkCategory() {
		return this.#legoPiece.bricklinkCategory;
	}

	get legoPiece() {
		return this.#legoPiece;
	}

	async save() {
		await database.saveLegoSetPiece(this);
	}

	getImagePath() {
		return this.#legoPiece.getImagePath();
	}

	getMinifigImagePath() {
		return this.#legoPiece.getMinifigImagePath();
	}
}

export class LegoSetTheme {
	/** @type {string} */
	#brand = "";
	/** @type {number} */
	databaseId;
	/** @type {string} */
	bricklinkId;
	/** @type {string} */
	bricklinkName;

	/**
	 * @param {number} databaseId
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 */
	constructor(databaseId, bricklinkId, bricklinkName) {
		this.databaseId = databaseId;
		this.bricklinkId = bricklinkId;
		this.bricklinkName = bricklinkName;
	}
}

export class LegoSet {
	/** @type {LegoPieces} */
	static legoPieces;

	/** @type {string} */
	#brand = "";
	/** @type {number} */
	databaseId;
	/** @type {string} */
	setNumber;
	/** @type {string} */
	name;
	/** @type {LegoSetTheme} */
	theme;
	/** @type {number} */
	releaseYear;
	/** @type {number} */
	pieceCount;
	/** @type {number} */
	minifigCount;
	/** @type {number} */
	legoSetCount;
	/** @type {LegoSetPieces}	*/
	#normalPieces;
	/** @type {LegoSetPieces} */
	#minifigs;
	/** @type {LegoSetPieces} */
	#extraPieces;
	/** @type {LegoSetPieces} */
	#counterpartPieces;

	/**
	 * @param {number} databaseId
	 * @param {string} setNumber
	 * @param {string} name
	 * @param {LegoSetTheme} theme
	 * @param {number} releaseYear
	 * @param {number} pieceCount
	 * @param {number} minifigCount
	 * @param {number} [legoSetCount]
	 * @param {LegoSetPieces} [normalPieces]
	 * @param {LegoSetPieces} [minifigs]
	 * @param {LegoSetPieces} [extraPieces]
	 * @param {LegoSetPieces} [counterpartPieces]
	 */
	constructor(
		databaseId,
		setNumber,
		name,
		theme,
		releaseYear,
		pieceCount,
		minifigCount,
		legoSetCount = 1,
		normalPieces = new LegoSetPieces("normal"),
		minifigs = new LegoSetPieces("minifig"),
		extraPieces = new LegoSetPieces("extra"),
		counterpartPieces = new LegoSetPieces("counterpart")
	) {
		this.databaseId = databaseId;
		this.setNumber = setNumber;
		this.name = name;
		this.theme = theme;
		this.releaseYear = releaseYear;
		this.pieceCount = pieceCount;
		this.minifigCount = minifigCount;
		this.legoSetCount = legoSetCount;
		this.#normalPieces = normalPieces;
		this.#minifigs = minifigs;
		this.#extraPieces = extraPieces;
		this.#counterpartPieces = counterpartPieces;
	}

	get normalPieces() {
		return this.#normalPieces;
	}

	get minifigs() {
		return this.#minifigs;
	}

	get extraPieces() {
		return this.#extraPieces;
	}

	get counterpartPieces() {
		return this.#counterpartPieces;
	}

	async save() {
		await database.saveLegoSet(this);
	}

	getImagePath() {
		return path.join(SET_IMAGES_PATH, `${this.setNumber}.png`);
	}

	/**
	 * @param {string} setNumber
	 */
	static getImagePath(setNumber) {
		return path.join(SET_IMAGES_PATH, `${setNumber}.png`);
	}

	/**
	 * @param {LegoSetPieces} legoSetPieces
	 * @param {LegoSetPieceInfo[]} newLegoSetPieces
	 */
	async addPieces(legoSetPieces, newLegoSetPieces) {
		for (let newLegoSetPiece of newLegoSetPieces) {
			let legoPieceId = LegoSet.legoPieces.getDatabaseId(
				newLegoSetPiece.bricklinkId,
				newLegoSetPiece.bricklinkName,
				newLegoSetPiece.color,
				newLegoSetPiece.bricklinkCategory
			);

			/** @type {LegoPiece | null} */
			let legoPiece = null;

			if (legoPieceId == null) {
				legoPiece = await LegoSet.legoPieces.add(
					newLegoSetPiece.bricklinkId,
					newLegoSetPiece.bricklinkName,
					newLegoSetPiece.color,
					newLegoSetPiece.bricklinkCategory
				);
			} else {
				legoPiece = /** @type {LegoPiece} */ (LegoSet.legoPieces.get(legoPieceId));
			}

			await legoSetPieces.add(this, legoPiece, newLegoSetPiece.amountNeeded, newLegoSetPiece.amountFound);
		}
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addNormalPieces(newSetPieces) {
		await this.addPieces(this.#normalPieces, newSetPieces);
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addMinifigs(newSetPieces) {
		await this.addPieces(this.#minifigs, newSetPieces);
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addExtraPieces(newSetPieces) {
		await this.addPieces(this.#extraPieces, newSetPieces);
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addCounterpartPieces(newSetPieces) {
		await this.addPieces(this.#counterpartPieces, newSetPieces);
	}
}
