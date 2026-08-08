// Imports
import { PGlite } from "@electric-sql/pglite";
import { eq, or } from "drizzle-orm";
import { drizzle } from "drizzle-orm/pglite";
import { migrate } from "drizzle-orm/pglite/migrator";
import { DATABASE_PATH, MIGRATIONS_PATH } from "../constants.js";
import { componentLegoPiecesDBTable, legoColorsDBTable, legoPiecesDBTable, legoSetPiecesDBTable, legoSetsDBTable, legoSetThemesDBTable, stickeredLegoPiecesDBTable } from "./schema.js";
/** @import { CompoundLegoPiece, LegoColor, LegoPiece, LegoSet, LegoSetPiece, StickeredLegoPiece } from "../models.js" */
/** @import { PgliteDatabase } from "drizzle-orm/pglite" */
/** @import { LegoSetPieceType, LegoSetStatus } from "../types.js" */


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
   let result = await database
      .insert(legoPiecesDBTable)
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
 * @param {StickeredLegoPiece} stickeredLegoPiece
 */
export async function addStickeredLegoPiece(stickeredLegoPiece) {
   await database
      .insert(stickeredLegoPiecesDBTable)
      .values({
         stickeredLegoPieceId: stickeredLegoPiece.databaseId,
         baseLegoPieceId: stickeredLegoPiece.baseLegoPiece.databaseId
      });
}

export async function getStickerRelations() {
   return await database
      .select({
         stickeredLegoPieceId: stickeredLegoPiecesDBTable.stickeredLegoPieceId,
         baseLegoPieceId: stickeredLegoPiecesDBTable.baseLegoPieceId
      }).from(stickeredLegoPiecesDBTable);
}

/**
 * @param {CompoundLegoPiece} compoundLegoPiece
 */
export async function addCompoundLegoPiece(compoundLegoPiece) {
   for (let componentLegoPiece of compoundLegoPiece.componentLegoPieces) {
      await database
         .insert(componentLegoPiecesDBTable)
         .values({
            componentLegoPieceId: componentLegoPiece.databaseId,
            compoundLegoPieceId: compoundLegoPiece.databaseId
         });
   }
}

export async function getComponentRelations() {
   return await database
      .select({
         componentLegoPieceId: componentLegoPiecesDBTable.componentLegoPieceId,
         compoundLegoPieceId: componentLegoPiecesDBTable.compoundLegoPieceId
      }).from(componentLegoPiecesDBTable);
}

/**
 * @param {number} legoSetId
 */
export async function getLegoSetPieces(legoSetId) {
   return await database
      .select()
      .from(legoSetPiecesDBTable)
      .where(eq(legoSetPiecesDBTable.legoSetId, legoSetId))
      .orderBy(legoSetPiecesDBTable.databaseId);
}

/**
 * @param {number} legoSetId
 * @param {number} legoPieceId
 * @param {LegoSetPieceType} legoSetPieceType
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
 * @param {LegoSetStatus} status
 */
export async function addLegoSet(setNumber, name, themeId, releaseYear, pieceCount, minifigCount, legoSetCount, status) {
   let result = await database.insert(legoSetsDBTable)
      .values({
         setNumber,
         name,
         themeId,
         releaseYear,
         pieceCount,
         minifigCount,
         legoSetCount,
         status
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
         legoSetCount: legoSet.legoSetCount,
         status: legoSet.status
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
   addStickeredLegoPiece,
   getStickerRelations,
   addCompoundLegoPiece,
   getComponentRelations,
   getLegoSetPieces,
   addLegoSetPiece,
   getLegoSets,
   addLegoSet,
   saveLegoSet,
   saveLegoSetPiece
};
