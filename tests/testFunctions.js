// Imports
import fs from "fs";
import path from "path";
import { TEST_DIRECTORY } from "../src/constants.js";


// Functions
/**
 * Gets the relative file path of a file in the test directory
 *
 * @param {string} fileName The name of the file
 * @returns {string}
 */
export function getRelativeFilePath(fileName) {
	return path.join(TEST_DIRECTORY, fileName);
}

/**
 * Reads a text file from the test directory (ex. txt, html)
 *
 * @param {string} fileName The name of the file
 * @returns {string} The contents of the file
 */
export function readTextFile(fileName) {
	return fs.readFileSync(getRelativeFilePath(fileName), "utf8");
}

/**
 * Reads an buffer file from the test directory (ex. images)
 *
 * @param {string} fileName The name of the file
 * @returns {ArrayBuffer} The contents of the file
 */
export function readBufferFile(fileName) {
	return fs.readFileSync(getRelativeFilePath(fileName));
}
