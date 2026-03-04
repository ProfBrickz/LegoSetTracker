import { IpcRendererEvent } from "electron";
import { LegoSetSearchResult, Theme } from "../types.js";

export interface ElectronAPI {
   loadPage(page: string, pageParams?: Object, params?: Object): void;
   setTheme(theme: Theme): void;
   searchLegoSets(): void;
   onSearchResults(callback: (event: IpcRendererEvent, searchResults: (LegoSetSearchResult & { theme: string, image: string; })[]) => void): void;
}
