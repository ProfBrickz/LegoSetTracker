// Imports
import { LegoSetPiece } from "./models.js";


// Types
/**
 * @typedef LegoSetInfo
 * @type {Object}
 * @property {string} setNumber The unique identifier for the set
 * @property {string} name The official name of the LEGO set
 * @property {string} theme The theme of the LEGO set
 * @property {number} releaseYear The release year
 * @property {number} pieceCount The total number of LEGO pieces in the set
 * @property {number} minifigCount The total number of minifigures in the set
*/

/**
 * @typedef LegoSetPieceInfo
 * @type {Object}
 * @property {LegoSetPiece[]} normalPieces The normal pieces in the set
 * @property {LegoSetPiece[]} minifigs The minifigures in the set
 * @property {LegoSetPiece[]} extraPieces The extra pieces in the set
 * @property {LegoSetPiece[]} counterparts The counterpart pieces in the set
 */
