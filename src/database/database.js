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

export default {
   getLegoColors
};
