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
export async function getLegoColors() {
   let colors = [];

   let results = await database.select().from(legoColorsDBTable);

   for (let result of results) {
      colors.push(new LegoColor(
         result.databaseId,
         result.bricklinkId,
         result.bricklinkName,
         result.legoId,
         result.legoName
      ));
   }

   return colors;
}

export default {
   getLegoColors
};
