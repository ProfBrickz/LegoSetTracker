import type { ElectronAPI } from "./electronAPI.d.ts";

export { };

declare global {
   interface Window {
      electronAPI: ElectronAPI;
   }

   interface DocumentEventMap {
      pageLoad: CustomEvent;
   }
}
