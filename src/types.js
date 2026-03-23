// Imports
import { LegoColor, LegoSetPiece } from "./models.js";


// Types
/**
 * @typedef {Object} LegoSetInfo
 * @property {string} setNumber The unique identifier for the set
 * @property {string} name The official name of the LEGO set
 * @property {string} themeId The id of the theme of the LEGO set
 * @property {number} releaseYear The release year
 * @property {number} pieceCount The total number of LEGO pieces in the set
 * @property {number} minifigCount The total number of minifigures in the set
*/

/**
 * @typedef {Object} LegoSetPieceInfo
 * @property {LegoSetPiece[]} normalPieces The normal pieces in the set
 * @property {LegoSetPiece[]} minifigs The minifigures in the set
 * @property {LegoSetPiece[]} extraPieces The extra pieces in the set
 * @property {LegoSetPiece[]} counterparts The counterpart pieces in the set
 */

/**
 * @typedef {Object} LegoSetSearchResult
 * @property {string} setNumber The unique identifier for the set
 * @property {string} name The official name of the LEGO set
 * @property {string} themeId The ID of the theme associated with the set (ex. 65.806.258)
 */

/**
 * @typedef {"light" | "dark" | "system"} Theme
 */

/**
 * @typedef {"normal" | "minifig" | "extra" | "counterpart"} LegoSetPieceType
 */

/**
 * @typedef {Object} LegoSetsTableRow
 * @property {string} image
 * @property {number} databaseId
 * @property {string} name
 * @property {string} setNumber
 * @property {string} theme
 * @property {number} releaseYear
 * @property {number} pieceCount
 * @property {number} minifigCount
 * @property {number} legoSetCount
 */

/**
 * @typedef {Object} LegoSetTableRow
 * @property {string} image
 * @property {number} pieceId
 * @property {number} amountFound
 * @property {number} amountLeft
 * @property {number} amountNeeded
 * @property {string} bricklinkName
 * @property {string} bricklinkId
 * @property {LegoColor | null} color
 * @property {string} bricklinkCategory
 */
