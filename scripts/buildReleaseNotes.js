// Imports
import fs from 'fs';


const PackageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
const version = PackageJson.version;

let releaseNotes = fs.readFileSync('scripts/RELEASE_NOTES-template.md', 'utf8');
releaseNotes = releaseNotes.replace(/@@VERSION@@/g, version);

fs.writeFileSync('RELEASE_NOTES.md', releaseNotes);
console.log(`Release notes for version ${version} generated in RELEASE_NOTES.md`);
