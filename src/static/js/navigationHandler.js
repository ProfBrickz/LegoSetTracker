// Imports
const fs = require("fs");


// Elements
const contentElement = document.getElementById("content");


// Functions
/**
 * Loads the specified page into the content element.
 * @param {string} page - The name of the page to load (without extension).
 * @returns {void}
 */
function loadPage(page) {
   let html = fs.readFileSync(`src/views/pages/${page}`, "utf8");


   if (contentElement) {
      contentElement.innerHTML = html;
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
