// Imports
import crypto from "crypto";
/** @import { BuildResult } from "electron-builder" */
import fs from "fs";
import path from "path";


// Constants
const BUILD_FOLDER = path.join(import.meta.dirname, "../builds");
const ARCHIVE_EXTENSIONS = [".7z", ".zip", ".tar.gz"];
const OPERATING_SYSTEMS = ["Windows", "Linux", "Mac"];
const ARCHITECTURES = ["x86_64", "x86_32", "arm64", "armv7l"];
const ARCHITECTURE_RENAME_MAP = {
	"x64": "x86_64",
	"amd64": "x86_64",
	"ia32": "x86_32",
	"aarch64": "arm64"
};


// Functions
/**
 * @param {Buffer} data
 * @param {string} algorithm
 */
function calculateChecksum(data, algorithm) {
	const hash = crypto.hash(algorithm, data);
	return hash;
}

/**
 * Moves a file to a new folder with updated naming conventions.
 *
 * @param {string} oldFilePath - The path of the file to move.
 */
function moveFile(oldFilePath) {
	// Skip nsis.7z files
	if (oldFilePath.endsWith("nsis.7z")) return;

	let fileName = path.basename(oldFilePath);

	// Make file names consistent
	for (let [oldArchitecture, newArchitecture] of Object.entries(ARCHITECTURE_RENAME_MAP)) {
		if (fileName.includes(oldArchitecture)) {
			fileName = fileName.replace(oldArchitecture, newArchitecture);
			break;
		}
	}

	// Remove " Installer" from the new path if it
	if (ARCHIVE_EXTENSIONS.some(extension => fileName.endsWith(extension))) {
		fileName = fileName.replace(" Installer", "");
	}

	// Find the folder to move the file to
	let newFilePath = path.join(BUILD_FOLDER, fileName);

	// Move the file to the new folder
	fs.renameSync(oldFilePath, newFilePath);

	// Generate SHA1 and SHA256 checksums
	let data = fs.readFileSync(newFilePath);
	let sha1Checksum = calculateChecksum(data, "sha1");
	let sha256Checksum = calculateChecksum(data, "sha256");

	fs.writeFileSync(`${newFilePath}.sha1`, sha1Checksum);
	fs.writeFileSync(`${newFilePath}.sha256`, sha256Checksum);

	return;
}

/**
 * Handles the completion of artifact builds by creating necessary folder structures and moving artifacts.
 *
 * @param {BuildResult} result The result object containing `artifactPaths`.
 */
async function artifactBuildCompleted(result) {
	// Create the build folder structure.
	if (!fs.existsSync(BUILD_FOLDER)) fs.mkdirSync(BUILD_FOLDER);
	for (let operatingSystem of OPERATING_SYSTEMS) {
		let operatingSystemFolder = path.join(BUILD_FOLDER, operatingSystem);
		if (!fs.existsSync(operatingSystemFolder)) fs.mkdirSync(operatingSystemFolder);

		for (let architecture of ARCHITECTURES) {
			let architectureFolder = path.join(operatingSystemFolder, architecture);
			if (!fs.existsSync(architectureFolder)) fs.mkdirSync(architectureFolder);
		}
	}

	// Move each artifact to its respective folder.
	for (let filePath of result.artifactPaths) {
		moveFile(filePath);
	}

	// Delete any empty folders
	for (let operatingSystem of OPERATING_SYSTEMS) {
		let operatingSystemFolder = path.join(BUILD_FOLDER, operatingSystem);

		for (let architecture of ARCHITECTURES) {
			let architectureFolder = path.join(operatingSystemFolder, architecture);

			if (fs.readdirSync(architectureFolder).length <= 0) fs.rmdirSync(architectureFolder);
		}

		if (fs.readdirSync(operatingSystemFolder).length <= 0) fs.rmdirSync(operatingSystemFolder);
	}
};


// Export the function as the default export
export default artifactBuildCompleted;
