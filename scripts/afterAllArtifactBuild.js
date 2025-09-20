// Imports
const fs = require("fs");
const path = require("path");


// Constants
const BUILD_FOLDER = path.join(__dirname, "../builds");
const ARCHIVE_EXTENSIONS = [".7z", ".zip", ".tar.gz"];
const OPERATING_SYSTEMS = ["Windows", "Linux", "Mac"];
const ARCHITECTURES = ["x86_64", "x86_32", "arm64", "armv7l"];
const ARCHITECTURE_RENAME_MAP = {
   "x64": "x86_64",
   "amd64": "x86_64",
   "ia32": "x86_32",
   "aarch64": "arm64"
};


/**
 * @param {string} oldFilePath - The path of the file
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

   // Get the operating system and architecture from the file name
   let operatingSystem = OPERATING_SYSTEMS.find(os => fileName.includes(os)) || "";
   let architecture = ARCHITECTURES.find(arch => fileName.includes(arch)) || "";
   let folder = path.join(BUILD_FOLDER, operatingSystem, architecture);

   // Find the folder to move the file to
   let newFilePath = path.join(folder, fileName);

   // Move the file to the new folder
   fs.renameSync(oldFilePath, newFilePath);

   return;
}

/**
 * Handles post-packaging file operations for Electron Builder
 * @param {import("electron-builder").BuildResult} result - The build result object from Electron Builder.
 * @returns {Promise<void>}
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
exports.default = artifactBuildCompleted;
