import type { ElectronAPI } from "./electronAPI.js";

export { };

declare global {
   interface Window {
      electronAPI: ElectronAPI;
   }

   interface DocumentEventMap {
      pageLoad: CustomEvent;
   }
}
