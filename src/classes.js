export class LegoColor {
	/** @type {number} */
	bricklinkId;
	/** @type {string} */
	bricklinkName;
	/** @type {number} */
	legoId;
	/** @type {string} */
	legoName;

	/**
	 * @param {number} bricklinkId
	 * @param {string} bricklinkName
	 * @param {number} legoId
	 * @param {string} legoName
	 * @returns {LegoColor}
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
	/** @type {LegoColor} */
	color;
	/** @type {string} */
	bricklinkCategory;

	/**
	 * @param {string} bricklinkId
	 * @param {string} bricklinkName
	 * @param {LegoColor} color
	 * @param {string} bricklinkCategory
	 * @returns {LegoPiece}
	 */
	constructor(bricklinkId, bricklinkName, color, bricklinkCategory) {
		this.bricklinkId = bricklinkId;
		this.bricklinkName = bricklinkName;
		this.color = color;
		this.bricklinkCategory = bricklinkCategory;
	}
}

export class LegoSetPiece {
	/**
	 * @private
	 * @type {LegoPiece}
	 */
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
	 * @returns {LegoSetPiece}
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
	/**
	 * @private
	 * @type {LegoPiece[]}
	 */
	#normalPieces;
	/**
	 * @private
	 * @type {LegoPiece[]}
	 */
	#minifigs;
	/**
	 * @private
	 * @type {LegoPiece[]}
	 */
	#extraPieces;
	/**
	 * @private
	 * @type {LegoPiece[]}
	 */
	#counterpartPieces;

	/**
	 * @param {string} setNumber
	 * @param {string} name
	 * @param {string} theme
	 * @param {number} releaseYear
	 * @param {number} pieceCount
	 * @param {number} minifigCount
	 * @param {number} [legoSetCount=1]
	 * @param {LegoPiece[]} [normalPieces=[]]
	 * @param {LegoPiece[]} [minifigs=[]]
	 * @param {LegoPiece[]} [extraPieces=[]]
	 * @param {LegoPiece[]} [counterpartPieces=[]]
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
