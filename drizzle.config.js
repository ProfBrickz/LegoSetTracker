// Imports
import { defineConfig } from "drizzle-kit";


export default defineConfig({
	out: "./drizzle",
	schema: "./src/database/schema.js",
	dialect: "postgresql",
	driver: "pglite",
	casing: "snake_case",
	breakpoints: false,
	dbCredentials: {
		url: "file:./database"
	}
});
