import "electron";
import { LegoSetSearchResult, Theme } from "../types.js";

export { };

declare global {
	namespace Electron {
		interface IpcMain {
			on<C extends keyof IpcEventMap>(
				channel: C,
				listener: (
					event: IpcMainEvent,
					...args: IpcEventMap[C]
				) => void
			): this;
			once<C extends keyof IpcEventMap>(
				channel: C,
				listener: (
					event: IpcMainEvent,
					...args: IpcEventMap[C]
				) => void
			): this;
			handle<C extends keyof IpcRequestsMap>(
				channel: C,
				listener: (
					event: IpcMainInvokeEvent,
					...args: IpcRequestsMap[C]["args"]
				) => Promise<IpcRequestsMap[C]["returns"]> | IpcRequestsMap[C]["returns"]
			): void;
			handleOnce<C extends keyof IpcRequestsMap>(
				channel: C,
				listener: (
					event: IpcMainInvokeEvent,
					...args: IpcRequestsMap[C]["args"]
				) => Promise<IpcRequestsMap[C]["returns"]> | IpcRequestsMap[C]["returns"]
			): void;
		}

		interface ipcRenderer {
			on<C extends keyof IpcEventMap>(
				channel: C,
				listener: (
					event: IpcRendererEvent,
					...args: IpcEventMap[C]
				) => void
			): this;
			once<C extends keyof IpcEventMap>(
				channel: C,
				listener: (
					event: IpcRendererEvent,
					...args: IpcEventMap[C]
				) => void
			): this;
			send<C extends keyof IpcEventMap>(
				channel: C,
				...args: IpcEventMap[C]
			): this;
			sendSync<C extends keyof IpcEventMap>(
				channel: C,
				...args: IpcEventMap[C]
			): this;
			invoke<C extends keyof IpcRequestsMap>(
				channel: C,
				...args: IpcRequestsMap[C]["args"]
			): this;
		}

		interface WebContents {
			send<C extends keyof IpcEventMap>(
				channel: C,
				...args: IpcEventMap[C]
			): void;
		}
	}
}

type IpcEventMap = {
	addSet: [
		setNumber: string
	];
	setTheme: [
		theme: Theme
	];
	themeChange: [
		theme: Theme
	];
	changeSetCount: [
		rowIndex: number,
		value: number
	];
};

type IpcRequestsMap = {
	loadPage: {
		args: [
			page: string,
			pageParams: Record<string, unknown>,
			params: Record<string, unknown>
		],
		returns: {
			page: string,
			html: string,
			params: Record<string, unknown>;
		};
	};
	searchLegoSets: {
		args: [
			searchQuery: string,
			options: {
				themeId: string,
				startYear: number,
				endYear: number;
			}
		];
		returns: (LegoSetSearchResult & { theme: string, image: string; })[];
	};
};
