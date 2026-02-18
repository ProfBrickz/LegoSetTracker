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
}
