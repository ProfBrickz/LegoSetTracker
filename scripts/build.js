// Imports
import { execSync } from "child_process";
import fs from "fs";
import path from "path";


// Types
/**
 * @typedef ContainerRuntime
 * @type {"podman" | "docker"}
 */


// Constants
let CONTAINER_NAME = "LegoSetTracker-builder";


// Functions
/**
 * Detects the container runtime to use.
 *
 * @returns {ContainerRuntime | null} The name of the container runtime to use, or `null` if neither
 */
function detectContainerRuntime() {
	try {
		execSync("podman --version", { stdio: "inherit" });

		return "podman";
	} catch (error) {
		try {
			execSync("docker --version", { stdio: "inherit" });

			return "docker";
		} catch (error) {
			return null;
		}
	}
}

/**
 * Checks if the container exists.
 *
 * @returns {boolean} `true` if the container exists, `false` if it does not.
 */
function checkContainerExists() {
	try {
		return execSync(
			`${containerRuntime} ps -a --format="{{.Names}}"`,
			{ encoding: "utf8", stdio: "inherit" }
		).split("\n").includes("LegoSetTracker-builder");
	} catch (error) {
		return false;
	}
}

/**
 * Creates the container.
 * 
 * @returns {void}
 */
function createContainer() {
	console.log("Creating LegoSetTracker-builder container...");
	execSync(`${containerRuntime} compose -f ${currentFolder}/container/compose.yml up -d`, { stdio: "inherit" });
}

/**
 * Manages the container.
 * If the container does not exist, it will be created.
 * If the container is not running, it will be started.
 * If the container is not set to this folder, it will be recreated.
 * 
 * @returns {void}
 */
function manageContainer() {
	let containerExist = checkContainerExists();

	if (!containerExist) {
		createContainer();
		return;
	}

	let containerFolder = execSync(
		`${containerRuntime} inspect ${CONTAINER_NAME} --format="{{range .Mounts}}{{.Source}}\n{{end}}"`,
		{ encoding: "utf8", stdio: "inherit" }
	).trim();

	if (currentFolder !== containerFolder) {
		execSync(`${containerRuntime} stop ${CONTAINER_NAME}`, { stdio: "inherit" });
		execSync(`${containerRuntime} rm -f ${CONTAINER_NAME}`, { stdio: "inherit" });
		console.log(`${CONTAINER_NAME} removed`);

		createContainer();
		return;
	}

	console.log("Starting LegoSetTracker-builder container...");
	execSync(`${containerRuntime} start ${CONTAINER_NAME}`, { stdio: "inherit" });
}


// Main
let currentFolder = import.meta.dirname;
if (currentFolder.endsWith("/scripts")) currentFolder = currentFolder.slice(0, -8);

// Detect container runtime
let containerRuntime = detectContainerRuntime();

if (containerRuntime) console.log(`Using ${containerRuntime} as container runtime`);
else {
	console.log("Neither Podman nor Docker found!");
	process.exit(1);
}

// Remove old build directories
let distPath = path.join(currentFolder, "dist");
if (fs.existsSync(distPath)) {
	console.log("Removing dist folder");
	fs.rmSync(distPath, { recursive: true, force: true });
}

let buildsPath = path.join(currentFolder, "builds");
if (fs.existsSync(buildsPath)) {
	console.log("Removing builds folder");
	fs.rmSync(buildsPath, { recursive: true, force: true });
}

// Start or create container
manageContainer();

// Ensure dependencies are installed
console.log("Checking and installing dependencies...");

try {
	execSync(`${containerRuntime} exec --env CI=true LegoSetTracker-builder pnpm install --frozen-lockfile`, { stdio: "inherit" });
} catch (error) {
	console.log("Failed to install dependencies!");
	process.exit(1);
}

// Run build
console.log("Running electron-builder...");
execSync(`${containerRuntime} exec --env CI=true LegoSetTracker-builder pnpm run electron-build`, { stdio: "inherit" });

// Stop container
execSync(`${containerRuntime} stop LegoSetTracker-builder`, { stdio: "inherit" });
