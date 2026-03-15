// Imports
import { BarChart, createIcons, House, Menu, Monitor, Moon, Settings, Sun, ToyBrick } from "lucide";
import "./pages/addLegoSet.js";
import "./pages/legoSets.js";


const iconsConfig = {
   icons: {
      Menu, House, ToyBrick, BarChart, Settings, Sun, Moon, Monitor
   }
};

// Initial icon creation
document.addEventListener("DOMContentLoaded", () => {
   window.electronAPI.loadPage("home");
   createIcons(iconsConfig);
});

// Recreate icons on page change
document.addEventListener("pageLoad", () => {
   createIcons(iconsConfig);
});
