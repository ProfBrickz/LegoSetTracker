// Types
/**
 * @typedef Color
 * @type {Object}
 * @property {number} bricklinkId The unique identifier for the color on BrickLink
 * @property {string} bricklinkName The color's name on BrickLink
 * @property {number} legoId The unique identifier for the color on LEGO's website
 * @property {string} legoName The color's name on LEGO's website
*/

/**
 * @typedef Piece
 * @type {Object}
 * @property {string} bricklinkId The unique identifier
 * @property {Color | null} color The color of the piece
 * @property {string} name The name of the piece
 * @property {string} category The type of LEGO part (e.g. "Brick", "Plate")
*/

/**
 * @typedef SetInfo
 * @type {Object}
 * @property {string} name The official name of the LEGO set
 * @property {string} setNumber The unique identifier for the set
 * @property {string} theme The theme of the LEGO set
 * @property {number} year The release year
 * @property {number} pieceCount The total number of LEGO pieces in the set
 * @property {number} minifigCount The total number of minifigures in the set
*/

/**
 * @typedef SetPieceType
 * @type {"normal" | "minifig" | "extra" | "counterpart"}
 */

/**
 * @typedef SetPiece
 * @type {Object}
 * @property {Piece} piece The LEGO piece that is part of the set
 * @property {SetPieceType} type
 * @property {number} amountNeeded The number of that piece needed to complete the set
*/

/**
 * @typedef LegoSet
 * @type {Object}
 * @property {string} setNumber The unique identifier for the set
 * @property {string} name The official name of the LEGO set
 * @property {string} theme The theme of the LEGO set
 * @property {number} year The release year of the set
 * @property {number} pieceCount The total number of LEGO pieces in the set
 * @property {number} minifigCount The total number of minifigures in the set
 * @property {SetPiece[]} normalPieces The normal pieces in the set
 * @property {SetPiece[]} minifigs The minifigures in the set
 * @property {SetPiece[]} extraPieces The extra pieces in the set
 * @property {SetPiece[]} counterparts The counterpart pieces in the set
*/

/**
 * @typedef VersionPiece
 * @type {Object}
 * @property {Piece} piece The LEGO piece that is part of the
 * @property {number} amountNeeded The number of that piece needed to complete the set
 * @property {number} amountFound The number of that piece already fou
 */


/**
 * @typedef SetVersion
 * @type {Object}
 * @property {number} id The unique identifier for the
 * @property {LegoSet} set The LEGO set that this version belongs to
 * @property {string} name The name of this
 * @property {VersionPiece[]} pieces An array of objects representing the pieces required for
 */

/**
 * @typedef SetPieceInfo
 * @type {Object}
 * @property {SetPiece[]} normalPieces The normal pieces in the set
 * @property {SetPiece[]} minifigs The minifigures in the set
 * @property {SetPiece[]} extraPieces The extra pieces in the set
 * @property {SetPiece[]} counterparts The counterpart pieces in the set
 */
