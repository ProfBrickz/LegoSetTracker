import { Theme } from "../types.js";

export interface ElectronAPI {
   loadPage(page: string, pageParams?: Object, params?: Object): void;
   setTheme(theme: Theme): void;
   searchLegoSets(): void;
}
