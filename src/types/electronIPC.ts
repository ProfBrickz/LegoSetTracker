import "electron";

export { };

declare global {
	namespace Electron {
		interface IpcMain {
			on<C extends keyof IpcChannels>(
				channel: C,
				listener: (
					event: IpcMainEvent,
					...args: IpcChannels[C]
				) => void
			): this;
			once<C extends keyof IpcChannels>(
				channel: C,
				listener: (
					event: IpcMainEvent,
					...args: IpcChannels[C]
				) => void
			): this;
			handle<C extends keyof IpcChannels>(
				channel: C,
				listener: (
					event: IpcMainInvokeEvent,
					...args: IpcChannels[C]
				) => Promise<any> | any
			): void;
			handleOnce<C extends keyof IpcChannels>(
				channel: C,
				listener: (
					event: IpcMainInvokeEvent,
					...args: IpcChannels[C]
				) => Promise<any> | any
			): void;
		}

		interface ipcRenderer {
			on<C extends keyof IpcChannels>(
				channel: C,
				listener: (
					event: IpcRendererEvent,
					...args: IpcChannels[C]
				) => void
			): this;
			once<C extends keyof IpcChannels>(
				channel: C,
				listener: (
					event: IpcRendererEvent,
					...args: IpcChannels[C]
				) => void
			): this;
			send<C extends keyof IpcChannels>(
				channel: C,
				...args: IpcChannels[C]
			): this;
			invoke<C extends keyof IpcChannels>(
				channel: C,
				...args: IpcChannels[C]
			): this;
			sendSync<C extends keyof IpcChannels>(
				channel: C,
				...args: IpcChannels[C]
			): this;
		}
	}
}

type IpcChannels = {
	loadPage: [
		page: string,
		pageParams: Record<string, unknown>,
		params: Record<string, unknown>
	];
};
