// Imports
import { PGlite } from "@electric-sql/pglite";
import { eq, or } from "drizzle-orm";
import { drizzle } from "drizzle-orm/pglite";
import { DATABASE_PATH } from "../constants.js";
import { LegoColor } from "../models.js";
import { legoColorsDBTable, legoPiecesDBTable, legoSetPiecesDBTable, legoSetsDBTable, legoSetThemesDBTable } from "./schema.js";


// Setup
const client = new PGlite(DATABASE_PATH);
const database = drizzle(client);


// Functions
export async function getLegoColors() {
   return await database.select().from(legoColorsDBTable);
}

/**
 * @param {number} bricklinkId
 * @param {string} bricklinkName
 * @param {number | null} legoId
 * @param {string | null} legoName
 */
export async function addLegoColor(bricklinkId, bricklinkName, legoId, legoName) {
   let result = await database.insert(legoColorsDBTable)
      .values({
         bricklinkId,
         bricklinkName,
         legoId,
         legoName
      })
      .returning({ databaseId: legoColorsDBTable.databaseId });

   if (!result || result.length === 0) {
      throw new Error("Insert failed: no ID returned");
   }

   return result[0].databaseId;
}

export async function getLegoSetThemes() {
   return await database.select({
      databaseId: legoSetThemesDBTable.databaseId,
      bricklinkId: legoSetThemesDBTable.bricklinkId,
      bricklinkName: legoSetThemesDBTable.bricklinkName
   }).from(legoSetThemesDBTable);
}

/**
 * @param {string} bricklinkId
 * @param {string} bricklinkName
 */
export async function addLegoSetTheme(bricklinkId, bricklinkName) {
   let existingRows = await database.select({
      databaseId: legoSetThemesDBTable.databaseId
   })
      .from(legoSetThemesDBTable)
      .where(
         or(
            eq(legoSetThemesDBTable.bricklinkId, bricklinkId),
            eq(legoSetThemesDBTable.bricklinkName, bricklinkName)
         )
      );

   if (existingRows.length > 0) return existingRows[0].databaseId;

   let result = await database.insert(legoSetThemesDBTable)
      .values({ bricklinkId, bricklinkName })
      .returning({
         databaseId: legoSetThemesDBTable.databaseId
      });

   if (!result || result.length === 0) {
      throw new Error("Insert failed: no ID returned");
   }

   return result[0].databaseId;
}

export async function getLegoPieces() {
   return await database.select().from(legoPiecesDBTable);
}

/**
    * @param {string} bricklinkId
    * @param {string} bricklinkName
    * @param {LegoColor | null} color
    * @param {string} bricklinkCategory
    */
export async function addLegoPiece(bricklinkId, bricklinkName, color, bricklinkCategory) {
   let result = await database.insert(legoPiecesDBTable)
      .values({
         bricklinkId,
         bricklinkName,
         colorId: color?.databaseId || null,
         bricklinkCategory
      }).returning({
         databaseId: legoPiecesDBTable.databaseId
      });

   if (!result || result.length === 0) {
      throw new Error("Insert failed: no ID returned");
   }

   return result[0].databaseId;
}

/**
 * @param {number} legoSetId
 */
export async function getLegoSetPieces(legoSetId) {
   return await database.select()
      .from(legoSetPiecesDBTable)
      .where(eq(legoSetPiecesDBTable.legoSetId, legoSetId));
}

/**
 * @param {number} legoSetId
 * @param {number} legoPieceId
 * @param {import("../types.js").LegoSetPieceType} legoSetPieceType
 * @param {number} amountNeeded
 * @param {number} amountFound
 */
export async function addLegoSetPiece(legoSetId, legoPieceId, legoSetPieceType, amountNeeded, amountFound) {
   let result = await database.insert(legoSetPiecesDBTable)
      .values({
         legoSetId,
         legoPieceId,
         legoSetPieceType,
         amountNeeded,
         amountFound
      }).returning({
         databaseId: legoSetPiecesDBTable.databaseId
      });

   if (!result || result.length === 0) {
      throw new Error("Insert failed: no ID returned");
   }

   return result[0].databaseId;
}

export async function getLegoSets() {
   return await database.select().from(legoSetsDBTable);
}

/**
 * @param {string} setNumber
 * @param {string} name
 * @param {number} themeId
 * @param {number} releaseYear
 * @param {number} pieceCount
 * @param {number} minifigCount
 * @param {number} legoSetCount
 */
export async function addLegoSet(setNumber, name, themeId, releaseYear, pieceCount, minifigCount, legoSetCount) {
   let result = await database.insert(legoSetsDBTable)
      .values({
         setNumber,
         name,
         themeId,
         releaseYear,
         pieceCount,
         minifigCount,
         legoSetCount
      })
      .returning({
         databaseId: legoSetsDBTable.databaseId
      });

   if (!result || result.length === 0) {
      throw new Error("Insert failed: no ID returned");
   }

   return result[0].databaseId;
}


export default {
   getLegoColors,
   addLegoColor,
   getLegoSetThemes,
   addLegoSetTheme,
   getLegoPieces,
   addLegoPiece,
   getLegoSetPieces,
   addLegoSetPiece,
   getLegoSets,
   addLegoSet
};
