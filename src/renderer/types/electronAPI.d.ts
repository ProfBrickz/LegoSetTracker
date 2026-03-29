import { LegoSetSearchResult, Theme } from "../types.js";

export interface ElectronAPI {
   loadPage(page: string, options?: { pageParams?: Record<string, unknown>, params?: Record<string, unknown>; }): void;
   setTheme(theme: Theme): void;
   searchLegoSets(callback: (searchResults: (LegoSetSearchResult & { theme: string, image: string; })[]) => void): void;
   addLegoSet(setNumber: string): void;
   deleteLegoSet(databaseId: number): Promise<void>;
   changeLegoSetCount(legoSetId: number, setCount: number): void;
   changeAmountFound(legoSetId: number, pieceId: number, amountFound: number): void;
}
