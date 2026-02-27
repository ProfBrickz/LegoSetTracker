// Imports
import path from "path";


// Constants
export const IS_DEV_MODE = process.env.NODE_ENV === "development" || true;
export const SRC_PATH = import.meta.dirname;
export const PRELOAD_FILE = path.join(SRC_PATH, "../preload/preload.js");
export const TEST_PATH = path.join(SRC_PATH, "../tests");
export const VIEWS_PATH = path.join(SRC_PATH, "renderer/views");
export const PAGES_PATH = path.resolve(VIEWS_PATH, "pages");
export const LAYOUTS_PATH = path.resolve(VIEWS_PATH, "layouts");
