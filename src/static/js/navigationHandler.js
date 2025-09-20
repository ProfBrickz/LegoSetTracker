// Imports
const fs = require("fs");
const { IS_DEV_MODE } = require("../constaints.js");


// Elements
let contentElement = document.getElementById("content");


// Variables
let pagesFolderPath;
if (IS_DEV_MODE) {
   pagesFolderPath = "src/views/pages/";
} else {
   pagesFolderPath = `${__dirname}/pages/`;
}


// Functions
/**
 * Loads the specified page into the content element.
 * @param {string} page - The name of the page to load (without extension).
 * @returns {void}
 */
function loadPage(page) {
   let filePath = pagesFolderPath + page;


   try {
      let html = fs.readFileSync(filePath, "utf8");

      if (contentElement) {
         contentElement.innerHTML = html;
      }
   } catch (error) {
      console.error(`Error loading page ${page}:`, error);
      if (contentElement) {
         contentElement.innerHTML = `<h1>Error</h1><p>Could not load page: ${page}</p>`;
      }
   }
}


// Add event listeners to navigation links
document.querySelectorAll("nav a").forEach(link => {
   link.addEventListener("click", function () {
      const page = this.getAttribute("data-page");
      loadPage(page);
   });
});

// Load the initial page
loadPage("home.html");
