// Imports
import fs from "fs";
import path from "path";
import { TEST_DIRECTORY } from "../src/constants.js";


// Functions
export function loadFixture(fileName) {
	return fs.readFileSync(path.join(TEST_DIRECTORY, fileName), "utf8");
}
