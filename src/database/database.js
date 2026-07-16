// Imports
import { PGlite } from "@electric-sql/pglite";
import { eq, or } from "drizzle-orm";
import { drizzle } from "drizzle-orm/pglite";
import { migrate } from "drizzle-orm/pglite/migrator";
import { DATABASE_PATH, MIGRATIONS_PATH } from "../constants.js";
import { legoColorsDBTable, legoPiecesDBTable, legoSetPiecesDBTable, legoSetsDBTable, legoSetThemesDBTable } from "./schema.js";
/** @import { LegoColor, LegoSet, LegoSetPiece } from "../models.js" */
/** @import { PgliteDatabase } from "drizzle-orm/pglite" */


/** @type {PgliteDatabase} */
let database;

// Functions
export async function init() {
   let client = await PGlite.create(DATABASE_PATH);
   database = drizzle(client);

   console.log("Migrating database...");
   await migrate(database, { migrationsFolder: MIGRATIONS_PATH });
   console.log("database migrated");
}

export async function getLegoColors() {
   return await database.select().from(legoColorsDBTable);
}

/**
 * @param {number} brickLinkId
 * @param {string} brickLinkName
 * @param {number | null} legoId
 * @param {string | null} legoName
 */
export async function addLegoColor(brickLinkId, brickLinkName, legoId, legoName) {
   let result = await database.insert(legoColorsDBTable)
      .values({
         brickLinkId,
         brickLinkName,
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
      brickLinkId: legoSetThemesDBTable.brickLinkId,
      brickLinkName: legoSetThemesDBTable.brickLinkName
   }).from(legoSetThemesDBTable);
}

/**
 * @param {string} brickLinkId
 * @param {string} brickLinkName
 */
export async function addLegoSetTheme(brickLinkId, brickLinkName) {
   let existingRows = await database.select({
      databaseId: legoSetThemesDBTable.databaseId
   })
      .from(legoSetThemesDBTable)
      .where(
         or(
            eq(legoSetThemesDBTable.brickLinkId, brickLinkId),
            eq(legoSetThemesDBTable.brickLinkName, brickLinkName)
         )
      );

   if (existingRows.length > 0) return existingRows[0].databaseId;

   let result = await database.insert(legoSetThemesDBTable)
      .values({ brickLinkId, brickLinkName })
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
    * @param {string} brickLinkId
    * @param {string} brickLinkName
    * @param {LegoColor | null} color
    * @param {string} brickLinkCategory
    */
export async function addLegoPiece(brickLinkId, brickLinkName, color, brickLinkCategory) {
   let result = await database.insert(legoPiecesDBTable)
      .values({
         brickLinkId,
         brickLinkName,
         colorId: color?.databaseId || null,
         brickLinkCategory
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

/**
 * @param {LegoSet} legoSet
 */
export async function saveLegoSet(legoSet) {
   await database.update(legoSetsDBTable)
      .set({
         setNumber: legoSet.setNumber,
         name: legoSet.name,
         themeId: legoSet.theme.databaseId,
         releaseYear: legoSet.releaseYear,
         pieceCount: legoSet.pieceCount,
         minifigCount: legoSet.minifigCount,
         legoSetCount: legoSet.legoSetCount
      })
      .where(eq(legoSetsDBTable.databaseId, legoSet.databaseId));
}

/**
 * @param {LegoSetPiece} legoSetPiece
 */
export async function saveLegoSetPiece(legoSetPiece) {
   await database.update(legoSetPiecesDBTable)
      .set({
         amountNeeded: legoSetPiece.amountNeeded,
         amountFound: legoSetPiece.amountFound
      })
      .where(eq(legoSetPiecesDBTable.databaseId, legoSetPiece.databaseId));
}


// Default export
export default {
   init,
   getLegoColors,
   addLegoColor,
   getLegoSetThemes,
   addLegoSetTheme,
   getLegoPieces,
   addLegoPiece,
   getLegoSetPieces,
   addLegoSetPiece,
   getLegoSets,
   addLegoSet,
   saveLegoSet,
   saveLegoSetPiece
};
