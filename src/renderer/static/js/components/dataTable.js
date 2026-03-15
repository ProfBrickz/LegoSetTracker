// Imports
import { createTable, getCoreRowModel } from "@tanstack/table-core";


// Types
/**
 * @typedef {Record<string, unknown>} TableRow
 */
/**
 * @typedef {import("@tanstack/table-core").ColumnDef<TableRow, any>} TableColumn
 */


// Web Component
export class DataTable extends HTMLElement {
   /** @type {import("@tanstack/table-core").Table<TableRow>} */
   table;
   /** @type {boolean} */
   hideOnEmpty = false;

   constructor() {
      super();

      this.table = createTable({
         columns: [],
         /** @type {TableRow[]} */
         data: [],
         getCoreRowModel: getCoreRowModel(),
         onStateChange: () => { },
         state: {
            columnPinning: {}
         },
         renderFallbackValue: null
      });
   }

   static observedAttributes = ["hide-on-empty"];

   /**
    * @param {string} name
    * @param {string} oldValue
    * @param {string} newValue
    */
   attributeChangedCallback(name, oldValue, newValue) {
      if (name == "hide-on-empty") this.hideOnEmpty = newValue != null && newValue !== "false";
   }

   get columns() {
      return this.table.options.columns;
   }

   get data() {
      return this.table.options.data;
   }

   /**
    * @param {TableColumn[]} columns
    */
   set columns(columns) {
      this.table.options.columns = columns;
   }

   /**
    * @param {TableRow[]} data
    */
   set data(data) {
      this.table.options.data = data;
   }

   /**
    * @param {import("@tanstack/table-core").Header<TableRow, unknown>} header
    */
   #createHeader(header) {
      let th = document.createElement("th");

      th.innerText = header.column.columnDef.header?.toString() || "";

      return th;
   }

   #createThead() {
      let thead = document.createElement("thead");
      let row = document.createElement("tr");
      thead.appendChild(row);

      for (let header of this.table.getFlatHeaders()) {
         row.appendChild(this.#createHeader(header));
      }

      return thead;
   }

   /**
    * @param {import("@tanstack/table-core").Cell<TableRow, unknown>} cell
    */
   #createCell(cell) {
      let td = document.createElement("td");

      let { type, onChange, min, max } = cell.column.columnDef.meta || {};

      if (onChange && (type === "string" || type === "number")) {
         let input = document.createElement("input");

         if (type === "string") {
            input.type = "text";
            input.onchange = (event) => {
               let element = /** @type {HTMLInputElement} */ (event.target);
               if (!element) return;

               let value = element.value;
               this.table.options.data[cell.row.index][cell.column.id] = value;

               if (onChange) {
                  onChange(cell.row.index, value);
               }
            };
         } else if (type === "number") {
            input.type = "number";
            input.min = String(min);
            input.max = String(max);
            input.onchange = (event) => {
               let element = /** @type {HTMLInputElement} */ (event.target);
               if (!element) return;

               let value = element.valueAsNumber;
               this.table.options.data[cell.row.index][cell.column.id] = value;

               if (cell.column.columnDef.meta?.onChange) {
                  cell.column.columnDef.meta.onChange(cell.row.index, value);
               }
            };
         }

         input.value = String(cell.getValue());
         td.appendChild(input);
      } else if (type === "image") {
         let imageElement = document.createElement("img");
         imageElement.src = String(cell.getValue());

         td.appendChild(imageElement);
      } else if (type === "function" && typeof cell.column.columnDef.cell === "function") {
         let result = cell.column.columnDef.cell(cell.getContext());
         td.appendChild(result);
      } else {
         td.innerText = String(cell.getValue());
      }

      return td;
   }

   /**
    * @param {import("@tanstack/table-core").Row<TableRow>} row
   */
   #createRow(row) {
      let tr = document.createElement("tr");

      for (let cell of row.getAllCells()) {
         tr.appendChild(this.#createCell(cell));
      }

      return tr;
   }

   #createTbody() {
      let tbody = document.createElement("tbody");

      for (let row of this.table.getRowModel().rows) {
         tbody.appendChild(this.#createRow(row));
      }

      return tbody;
   }

   renderTable() {
      this.innerHTML = "";

      if (this.hideOnEmpty && this.table.options.data.length <= 0) {
         this.style.display = "none";
      } else {
         this.style.display = "";
      }

      let tableElement = document.createElement("table");
      tableElement.appendChild(this.#createThead());
      tableElement.appendChild(this.#createTbody());

      this.appendChild(tableElement);
   }
}

customElements.define("data-table", DataTable);
