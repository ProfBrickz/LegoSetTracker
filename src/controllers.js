// Imports
import { ClassMap } from "./classes.js";
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
