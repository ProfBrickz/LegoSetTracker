// Imports
import fs from "fs";
import path from "path";
import { TEST_DIRECTORY } from "../src/constants.js";


// Functions
/**
 * Loads a fixture file from the test directory.
 * @param {string} fileName - The name of the fixture file.
 * @param {"utf8" | null} [encoding="utf8"] - The file encoding.
 * @returns {string | Buffer} The contents of the fixture file.
 */
export function loadFixture(fileName, encoding = "utf8") {
	return fs.readFileSync(path.join(TEST_DIRECTORY, fileName), encoding);
}
