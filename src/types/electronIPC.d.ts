import "electron";
import { IpcMain, IpcRenderer } from "electron";
import { LegoSetSearchResult, Theme } from "../types.js";

export { };

export type TypedIpcMain = Omit<IpcMain, "on" | "once" | "handle" | "handleOnce"> & {
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
};

export type TypedIpcRenderer = Omit<IpcRenderer, "on" | "once" | "send" | "sendSync" | "invoke"> & {
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
	invoke<C extends string & keyof IpcRequestsMap>(
		channel: C,
		...args: IpcRequestsMap[C]["args"]
	): Promise<IpcRequestsMap[C]["returns"]>;
};

declare global {
	namespace Electron {
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
