// Imports
import { PGlite } from "@electric-sql/pglite";
import { drizzle } from "drizzle-orm/pglite";
import { DATABASE_PATH } from "../constants.js";
import { LegoColor } from "../models.js";
import { legoColorsDBTable } from "./schema.js";


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

export default {
   getLegoColors,
   addLegoColor
};
