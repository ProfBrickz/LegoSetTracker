// Imports
import path from "path";
import { PIECE_IMAGES_PATH, SET_IMAGES_PATH } from "./constants.js";
import { LegoPieces, LegoSetPieces } from "./controllers.js";


// Classes
export class LegoColor {
	/** @type {string} */
	#brand = "";
	/** @type {number | null} */
	databaseId;
	/** @type {number} */
	bricklinkId;
	/** @type {string} */
	bricklinkName;
	/** @type {number | null} */
	legoId;
	/** @type {string} */
	legoName;

	/**
	 * @param {number | null} databaseId
	 * @param {number} bricklinkId
	 * @param {string} bricklinkName
	 * @param {number | null} legoId
	 * @param {string} legoName
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
	/** @type {number | null} */
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
	 * @param {number | null} databaseId
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

		return path.join(PIECE_IMAGES_PATH, colorId.toString(), `${this.bricklinkId}.jpg`);
	}
}

export class LegoSetPiece {
	/** @type {string} */
	#brand = "";
	/** @type {number | null} */
	databaseId;
	/** @type {LegoPiece} */
	#legoPiece;
	/** @type {number} */
	amountNeeded;
	/** @type {number} */
	amountFound;

	// Make the constructor with the jsdoc string
	/**
	 * @param {number | null} databaseId
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

	getImagePath() {
		return this.#legoPiece.getImagePath();
	}
}

export class LegoSet {
	/** @type {LegoPieces} */
	static legoPieces;

	/** @type {string} */
	#brand = "";
	/** @type {number | null} */
	databaseId;
	/** @type {string} */
	setNumber;
	/** @type {string} */
	name;
	/** @type {string} */
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
	 * @param {number | null} databaseId
	 * @param {string} setNumber
	 * @param {string} name
	 * @param {string} theme
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
		normalPieces = new LegoSetPieces(),
		minifigs = new LegoSetPieces(),
		extraPieces = new LegoSetPieces(),
		counterpartPieces = new LegoSetPieces()
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

	getImagePath() {
		return path.join(SET_IMAGES_PATH, `${this.setNumber}.jpg`);
	}

	/**
	 * @param {string} setNumber
	 */
	static getImagePath(setNumber) {
		return path.join(SET_IMAGES_PATH, `${setNumber}.jpg`);
	}

	/**
	 * @param {LegoSetPieces} pieces
	 * @param {LegoSetPiece[]} newSetPieces
	 */
	addPieces(pieces, newSetPieces) {
		for (let newSetPiece of newSetPieces) {
			let legoPieceId = LegoSet.legoPieces.getDatabaseId(
				newSetPiece.bricklinkId,
				newSetPiece.bricklinkName,
				newSetPiece.color,
				newSetPiece.bricklinkCategory
			);
			/** @type {LegoPiece | null} */
			let legoPiece = null;

			if (legoPieceId == null) {
				legoPieceId = LegoSet.legoPieces.size;

				legoPiece = new LegoPiece(
					legoPieceId,
					newSetPiece.bricklinkId,
					newSetPiece.bricklinkName,
					newSetPiece.color,
					newSetPiece.bricklinkCategory
				);

				LegoSet.legoPieces.set(legoPieceId, legoPiece);
			} else {
				legoPiece = /** @type {LegoPiece} */ (LegoSet.legoPieces.get(legoPieceId));
			}

			pieces.set(pieces.size, new LegoSetPiece(
				pieces.size,
				legoPiece,
				newSetPiece.amountNeeded,
				newSetPiece.amountFound
			));
		}
	}

	/**
	 * @param {LegoSetPiece[]} newSetPieces
	 */
	addNormalPieces(newSetPieces) {
		this.addPieces(this.#normalPieces, newSetPieces);
	}

	/**
	 * @param {LegoSetPiece[]} newSetPieces
	 */
	addMinifigs(newSetPieces) {
		this.addPieces(this.#minifigs, newSetPieces);
	}

	/**
	 * @param {LegoSetPiece[]} newSetPieces
	 */
	addExtraPieces(newSetPieces) {
		this.addPieces(this.#extraPieces, newSetPieces);
	}

	/**
	 * @param {LegoSetPiece[]} newSetPieces
	 */
	addCounterpartPieces(newSetPieces) {
		this.addPieces(this.#counterpartPieces, newSetPieces);
	}
}
