// Imports
import fs from "fs";


const PackageJson = JSON.parse(fs.readFileSync("package.json", "utf8"));
const version = PackageJson.version;

let notes = fs.readFileSync("release-notes.md", "utf8");
if (!notes.endsWith("\n")) notes += "\n";

let releaseNotes = fs.readFileSync("scripts/RELEASE_NOTES-template.md", "utf8");
releaseNotes = releaseNotes.replace(/@@VERSION@@/g, version);
releaseNotes = releaseNotes.replace(/@@NOTES@@/g, notes);

fs.writeFileSync("RELEASE_NOTES.md", releaseNotes);
console.log(`Release notes for version ${version} generated in RELEASE_NOTES.md`);
