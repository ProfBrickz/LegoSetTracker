// Imports
import path from "path";


// Constants
export const IS_DEV_MODE = process.env.NODE_ENV === "development";

export const SRC_PATH = import.meta.dirname;
export const PRELOAD_FILE = path.join(SRC_PATH, "../preload/preload.js");
export const TEST_PATH = path.join(SRC_PATH, "../tests");
export const VIEWS_PATH = path.join(SRC_PATH, "../renderer/views");
export const PAGES_PATH = path.join(VIEWS_PATH, "pages");
export const LAYOUTS_PATH = path.join(VIEWS_PATH, "layouts");
export const MIGRATIONS_PATH = path.join(SRC_PATH, "../../drizzle");

let tempDataPath;
if (typeof process.versions.electron === "string" && process.versions.electron !== "" && !IS_DEV_MODE) {
	const { app } = await import("electron");
	tempDataPath = app.getPath("userData");
} else {
	tempDataPath = path.resolve(SRC_PATH, "../../data");
}

export const DATA_PATH = tempDataPath;
export const DATABASE_PATH = path.join(DATA_PATH, "database");
export const IMAGES_PATH = path.join(DATA_PATH, "images");
export const SET_IMAGES_PATH = path.join(IMAGES_PATH, "sets");
export const PIECE_IMAGES_PATH = path.join(IMAGES_PATH, "pieces");
export const MINIFIG_IMAGES_PATH = path.join(IMAGES_PATH, "minifigs");
