// Imports
import path from "path";
import { TypedClass } from "./classes.js";
import { MINIFIG_IMAGES_PATH, PIECE_IMAGES_PATH, SET_IMAGES_PATH } from "./constants.js";
import database from "./database/database.js";
import { LegoPieces, LegoSetPieces } from "./dataMaps.js";
/** @import { LegoSetPieceInfo } from "./types.js" */


// Classes
export class LegoColor extends TypedClass {
	/** @type {number} */
	databaseId;
	/** @type {number} */
	brickLinkId;
	/** @type {string} */
	brickLinkName;
	/** @type {number | null} */
	legoId;
	/** @type {string | null} */
	legoName;

	/**
	 * @param {number} databaseId
	 * @param {number} brickLinkId
	 * @param {string} brickLinkName
	 * @param {number | null} legoId
	 * @param {string | null} legoName
	 */
	constructor(databaseId, brickLinkId, brickLinkName, legoId, legoName) {
		super();

		this.databaseId = databaseId;
		this.brickLinkId = brickLinkId;
		this.brickLinkName = brickLinkName;
		this.legoId = legoId;
		this.legoName = legoName;
	}
}

export class LegoPiece extends TypedClass {
	/** @type {number} */
	databaseId;
	/** @type {string} */
	brickLinkId;
	/** @type {string} */
	brickLinkName;
	/** @type {LegoColor | null} */
	color;
	/** @type {string} */
	brickLinkCategory;

	/**
	 * @param {number} databaseId
	 * @param {string} brickLinkId
	 * @param {string} brickLinkName
	 * @param {LegoColor | null} color
	 * @param {string} brickLinkCategory
	*/
	constructor(databaseId, brickLinkId, brickLinkName, color, brickLinkCategory) {
		super();

		this.databaseId = databaseId;
		this.brickLinkId = brickLinkId;
		this.brickLinkName = brickLinkName;
		this.color = color;
		this.brickLinkCategory = brickLinkCategory;
	}

	getImagePath() {
		let colorId = this.color?.brickLinkId || 0;

		return path.join(PIECE_IMAGES_PATH, colorId.toString(), `${this.brickLinkId}.png`);
	}

	getMinifigImagePath() {
		return path.join(MINIFIG_IMAGES_PATH, `${this.brickLinkId}.png`);
	}
}

export class StickeredLegoPiece extends LegoPiece {
	/** @type {LegoPiece} */
	baseLegoPiece;

	/**
	 * @param {number} databaseId
	 * @param {string} brickLinkId
	 * @param {string} brickLinkName
	 * @param {LegoColor | null} color
	 * @param {string} brickLinkCategory
	 * @param {LegoPiece} baseLegoPiece
	*/
	constructor(databaseId, brickLinkId, brickLinkName, color, brickLinkCategory, baseLegoPiece) {
		super(databaseId, brickLinkId, brickLinkName, color, brickLinkCategory);

		this.baseLegoPiece = baseLegoPiece;
	}
}

export class CompoundLegoPiece extends LegoPiece {
	/** @type {LegoPiece[]} */
	components;

	/**
	 * @param {number} databaseId
	 * @param {string} brickLinkId
	 * @param {string} brickLinkName
	 * @param {LegoColor | null} color
	 * @param {string} brickLinkCategory
	 * @param {LegoPiece[]} components
	 */
	constructor(databaseId, brickLinkId, brickLinkName, color, brickLinkCategory, components = []) {
		super(databaseId, brickLinkId, brickLinkName, color, brickLinkCategory);

		this.components = components;
	}
}

export class LegoSetPiece extends TypedClass {
	/** @type {number} */
	databaseId;
	/**
	 * @readonly
	 * @type {LegoPiece}
	 */
	legoPiece;
	/** @type {number} */
	amountNeeded;
	/** @type {number} */
	amountFound;
	/** @type {StickeredLegoSetPiece[]} */
	stickeredLegoSetPieces = [];
	/** @type {CompoundLegoSetPiece[]} */
	compoundLegoSetPieces = [];

	// Make the constructor with the jsdoc string
	/**
	 * @param {number} databaseId
	 * @param {LegoPiece} legoPiece
	 * @param {number} amountNeeded
	 * @param {number} [amountFound=0]
	 */
	constructor(databaseId, legoPiece, amountNeeded, amountFound = 0) {
		super();

		this.databaseId = databaseId;
		this.legoPiece = legoPiece;
		this.amountNeeded = amountNeeded;
		this.amountFound = amountFound;
	}

	get brickLinkId() {
		return this.legoPiece.brickLinkId;
	}

	get brickLinkName() {
		return this.legoPiece.brickLinkName;
	}

	get color() {
		return this.legoPiece.color;
	}

	get brickLinkCategory() {
		return this.legoPiece.brickLinkCategory;
	}

	async save() {
		await database.saveLegoSetPiece(this);
	}

	getImagePath() {
		return this.legoPiece.getImagePath();
	}

	getMinifigImagePath() {
		return this.legoPiece.getMinifigImagePath();
	}

	getTotalStickeredPiecesFound() {
		let total = 0;

		for (let stickeredLegoSetPiece of this.stickeredLegoSetPieces) {
			total += stickeredLegoSetPiece.amountFound;
		}

		for (let compoundLegoSetPiece of this.compoundLegoSetPieces) {
			total += compoundLegoSetPiece.amountFound;
		}

		return total;
	}

	/**
	 * @param {number} amountFound
	 */
	syncAmountFound(amountFound) {
		let totalStickeredPiecesFound = this.getTotalStickeredPiecesFound();
		if (amountFound < totalStickeredPiecesFound) {
			amountFound = totalStickeredPiecesFound;
		}

		this.amountFound = amountFound;
	}
}

export class StickeredLegoSetPiece extends LegoSetPiece {
	/** @type {LegoSetPiece} */
	baseLegoSetPiece;

	/**
	 * @param {number} databaseId
	 * @param {StickeredLegoPiece} legoPiece
	 * @param {LegoSetPiece} baseLegoSetPiece
	 * @param {number} amountNeeded
	 * @param {number} [amountFound=0]
	 */
	constructor(databaseId, legoPiece, baseLegoSetPiece, amountNeeded, amountFound = 0) {
		super(databaseId, legoPiece, amountNeeded, amountFound);

		this.baseLegoSetPiece = baseLegoSetPiece;
	}

	/**
	 * @param {number} amountFound
	 */
	syncAmountFound(amountFound) {
		this.baseLegoSetPiece.amountFound += amountFound - this.amountFound;

		this.amountFound = amountFound;
	}
}

export class CompoundLegoSetPiece extends LegoSetPiece {
	/** @type {LegoSetPiece[]} */
	legoSetComponents;

	/**
	 * @param {number} databaseId
	 * @param {CompoundLegoPiece} legoPiece
	 * @param {number} amountNeeded
	 * @param {number} amountFound
	 * @param {LegoSetPiece[]} legoSetComponents
	 */
	constructor(databaseId, legoPiece, amountNeeded, amountFound = 0, legoSetComponents = []) {
		super(databaseId, legoPiece, amountNeeded, amountFound);

		this.legoSetComponents = legoSetComponents;
	}

	/**
	 * @param {number} amountFound
	 */
	syncAmountFound(amountFound) {
		let amountFoundDifference = amountFound - this.amountFound;

		for (let legoSetComponent of this.legoSetComponents) {
			legoSetComponent.amountFound += amountFoundDifference;
		}

		this.amountFound = amountFound;
	}

	updateAmountFound() {
		/** @type {number | null} */
		let minAmountFound = null;

		for (let legoSetComponent of this.legoSetComponents) {
			if (minAmountFound == null) {
				minAmountFound = legoSetComponent.amountFound;
				continue;
			}

			if (legoSetComponent.amountFound < minAmountFound) minAmountFound = legoSetComponent.amountFound;
		}

		if (minAmountFound == null) minAmountFound = 0;

		this.amountFound = minAmountFound;
	}
}

export class LegoSetTheme extends TypedClass {
	/** @type {number} */
	databaseId;
	/** @type {string} */
	brickLinkId;
	/** @type {string} */
	brickLinkName;

	/**
	 * @param {number} databaseId
	 * @param {string} brickLinkId
	 * @param {string} brickLinkName
	 */
	constructor(databaseId, brickLinkId, brickLinkName) {
		super();

		this.databaseId = databaseId;
		this.brickLinkId = brickLinkId;
		this.brickLinkName = brickLinkName;
	}
}

export class LegoSet extends TypedClass {
	/** @type {LegoPieces} */
	static legoPieces;

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
	/**
	 * @readonly
	 * @type {LegoSetPieces}
	 */
	normalPieces;
	/**
	 * @readonly
	 * @type {LegoSetPieces}
	 */
	minifigs;
	/**
	 * @readonly
	 * @type {LegoSetPieces}
	 */
	extraPieces;
	/**
	 * @readonly
	 * @type {LegoSetPieces}
	 */
	counterpartPieces;

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
		super();

		this.databaseId = databaseId;
		this.setNumber = setNumber;
		this.name = name;
		this.theme = theme;
		this.releaseYear = releaseYear;
		this.pieceCount = pieceCount;
		this.minifigCount = minifigCount;
		this.legoSetCount = legoSetCount;
		this.normalPieces = normalPieces;
		this.minifigs = minifigs;
		this.extraPieces = extraPieces;
		this.counterpartPieces = counterpartPieces;
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
				newLegoSetPiece.brickLinkId,
				newLegoSetPiece.brickLinkName,
				newLegoSetPiece.color,
				newLegoSetPiece.brickLinkCategory
			);

			/** @type {LegoPiece | null} */
			let legoPiece = null;

			if (legoPieceId == null) {
				legoPiece = await LegoSet.legoPieces.add(
					newLegoSetPiece.brickLinkId,
					newLegoSetPiece.brickLinkName,
					newLegoSetPiece.color,
					newLegoSetPiece.brickLinkCategory
				);
				if (legoPiece == null) return;
			} else {
				legoPiece = /** @type {LegoPiece} */ (LegoSet.legoPieces.get(legoPieceId));
			}

			await legoSetPieces.add(
				this,
				legoPiece,
				newLegoSetPiece.amountNeeded,
				newLegoSetPiece.amountFound
			);
		}
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addNormalPieces(newSetPieces) {
		await this.addPieces(this.normalPieces, newSetPieces);
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addMinifigs(newSetPieces) {
		await this.addPieces(this.minifigs, newSetPieces);
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addExtraPieces(newSetPieces) {
		await this.addPieces(this.extraPieces, newSetPieces);
	}

	/**
	 * @param {LegoSetPieceInfo[]} newSetPieces
	 */
	async addCounterpartPieces(newSetPieces) {
		await this.addPieces(this.counterpartPieces, newSetPieces);
	}

	/**
	 * @param {StickeredLegoPiece} stickeredLegoPiece
	 */
	getBaseSetPiece(stickeredLegoPiece) {
		return this.normalPieces.getBaseSetPiece(stickeredLegoPiece);
	}
}
