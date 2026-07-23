// Imports
import {
   BarChart,
   Check,
   ChevronDown,
   ChevronRight,
   createIcons,
   House,
   Menu, Monitor,
   Moon,
   Plus,
   Settings,
   Sun,
   ToyBrick,
   X
} from "lucide";
import "./pages/addLegoSet.js";
import "./pages/legoSet.js";
import "./pages/legoSets.js";


const iconsConfig = {
   icons: {
      BarChart,
      Check,
      ChevronDown,
      ChevronRight,
      House,
      Menu, Monitor,
      Moon,
      Plus,
      Settings,
      Sun,
      ToyBrick,
      X
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
