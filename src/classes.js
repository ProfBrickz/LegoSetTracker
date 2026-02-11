export class LegoColor {
	/** @type {number} */
	bricklinkId;
	/** @type {string} */
	bricklinkName;
	/** @type {number | null} */
	legoId;
	/** @type {string} */
	legoName;

	/**
	 * @param {number} bricklinkId
	 * @param {string} bricklinkName
	 * @param {number | null} legoId
	 * @param {string} legoName
	 */
	constructor(bricklinkId, bricklinkName, legoId, legoName) {
		this.bricklinkId = bricklinkId;
		this.bricklinkName = bricklinkName;
		this.legoId = legoId;
		this.legoName = legoName;
	}
}

export class LegoPiece {
	/** @type {string} */
	bricklinkId;
	/** @type {string} */
	bricklinkName;
	/** @type {LegoColor | null} */
	color;
	/** @type {string} */
	bricklinkCategory;

	/**
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 * @param {LegoColor | null} color
	 * @param {string} bricklinkCategory
	 */
	constructor(bricklinkId, bricklinkName, color, bricklinkCategory) {
		this.bricklinkId = bricklinkId;
		this.bricklinkName = bricklinkName;
		this.color = color;
		this.bricklinkCategory = bricklinkCategory;
	}
}

export class LegoSetPiece {
	/** @type {LegoPiece} */
	#legoPiece;
	/** @type {number} */
	amountNeeded;
	/** @type {number} */
	amountFound;

	// Make the constructor with the jsdoc string
	/**
	 * @param {LegoPiece} legoPiece
	 * @param {number} amountNeeded
	 * @param {number} [amountFound=0]
	 */
	constructor(legoPiece, amountNeeded, amountFound = 0) {
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
}

export class LegoSet {
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
	/** @type {LegoSetPiece[]}	*/
	#normalPieces;
	/** @type {LegoSetPiece[]} */
	#minifigs;
	/** @type {LegoSetPiece[]} */
	#extraPieces;
	/** @type {LegoSetPiece[]} */
	#counterpartPieces;

	/**
	 * @param {string} setNumber
	 * @param {string} name
	 * @param {string} theme
	 * @param {number} releaseYear
	 * @param {number} pieceCount
	 * @param {number} minifigCount
	 * @param {number} [legoSetCount=1]
	 * @param {LegoSetPiece[]} [normalPieces=[]]
	 * @param {LegoSetPiece[]} [minifigs=[]]
	 * @param {LegoSetPiece[]} [extraPieces=[]]
	 * @param {LegoSetPiece[]} [counterpartPieces=[]]
	 */
	constructor(
		setNumber,
		name,
		theme,
		releaseYear,
		pieceCount,
		minifigCount,
		legoSetCount = 1,
		normalPieces = [],
		minifigs = [],
		extraPieces = [],
		counterpartPieces = []
	) {
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
}
