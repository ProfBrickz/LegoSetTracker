// Imports
import path from "path";


// Constants
export const IS_DEV_MODE = process.env.NODE_ENV === "development";
export const SRC_DIRECTORY = import.meta.dirname;
export const TEST_DIRECTORY = path.join(SRC_DIRECTORY, "../tests");
export const VIEWS_PATH = path.join(SRC_DIRECTORY, "views");
export const PAGES_PATH = path.join(VIEWS_PATH, "pages");
export const LAYOUTS_PATH = path.join(VIEWS_PATH, "layouts");
