import { PGlite } from "@electric-sql/pglite";
import { drizzle } from "drizzle-orm/pglite";
import { DATABASE_PATH } from "../constants.js";


const client = new PGlite(DATABASE_PATH);
const database = drizzle(client);

export default database;
