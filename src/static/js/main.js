import { BarChart, createIcons, House, Menu, Monitor, Moon, Settings, Sun, ToyBrick } from "../../../node_modules/lucide/dist/esm/lucide.js";

const iconsConfig = {
   icons: {
      Menu, House, ToyBrick, BarChart, Settings, Sun, Moon, Monitor
   }
};

// Initial icon creation
document.addEventListener("DOMContentLoaded", () => {
   electronAPI.loadPage("home");
   createIcons(iconsConfig);
});

// Recreate icons on page change
document.addEventListener("pageChanged", () => {
   createIcons(iconsConfig);
});
