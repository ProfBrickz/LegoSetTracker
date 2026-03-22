// Imports
import { PGlite } from "@electric-sql/pglite";
import { drizzle } from "drizzle-orm/pglite";
import { DATABASE_PATH } from "../constants.js";
import { LegoColor } from "../models.js";
import { legoColorsDBTable, legoSetThemesDBTable } from "./schema.js";


// Setup
const client = new PGlite(DATABASE_PATH);
const database = drizzle(client);


// Functions
/**
 * @returns {Promise<Omit<LegoColor, "#brand">[]>}
 */
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
      bricklinkId: legoSetThemesDBTable.bricklinkId,
      bricklinkName: legoSetThemesDBTable.bricklinkName
   }).from(legoSetThemesDBTable);
}

/**
 * @param {string} bricklinkId
 * @param {string} bricklinkName
 */
export async function addLegoSetTheme(bricklinkId, bricklinkName) {
   return await database.insert(legoSetThemesDBTable)
      .values({
         bricklinkId,
         bricklinkName
      })
      .returning({
         databaseId: legoSetThemesDBTable.databaseId
      });
}


export default {
   getLegoColors,
   addLegoColor,
   getLegoSetThemes,
   addLegoSetTheme
};
