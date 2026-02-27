import { Theme } from "../types.js";

export interface ElectronAPI {
   loadPage(page: string, params?: Object): void;
   setTheme(theme: Theme): void;
   searchLegoSets(): void;
}
