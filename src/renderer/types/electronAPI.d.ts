import { LegoSetSearchResult, Theme } from "../types.js";

export interface ElectronAPI {
   loadPage(page: string, pageParams?: Object, params?: Object): void;
   setTheme(theme: Theme): void;
   searchLegoSets(callback: (searchResults: (LegoSetSearchResult & { theme: string, image: string; })[]) => void): void;
   addSet(setNumber: string): void;
   changeSetCount(setId: number, value: number): void;
}
