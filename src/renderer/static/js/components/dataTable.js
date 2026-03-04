// Imports
import { createTable, getCoreRowModel } from "@tanstack/table-core";


// Types
/**
 * @typedef {Record<string, unknown>} TableRow
 */
/**
 * @typedef {import("@tanstack/table-core").ColumnDef<TableRow, any>} TableColumn
 */


export class DataTable extends HTMLElement {
   /** @type {import("@tanstack/table-core").Table<TableRow>} */
   table;

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

      let { type, editable, onChange } = cell.column.columnDef.meta || {};

      if (editable && type) {
         let input = document.createElement("input");

         if (type === "string") {
            input.type = "text";
            input.onchange = (event) => {
               let element = /** @type {HTMLInputElement} */ (event.target);
               if (!element) return;

               let value = element.value;
               this.table.options.data[cell.row.index][cell.column.id] = value;

               if (onChange) {
                  onChange(value);
               }

               this.#handleChange(cell.column.id, cell.row.index, value);
            };
         } else if (type === "number") {
            input.type = "number";
            input.onchange = (event) => {
               let element = /** @type {HTMLInputElement} */ (event.target);
               if (!element) return;

               let value = Number(element.value);
               this.table.options.data[cell.row.index][cell.column.id] = value;

               this.#handleChange(cell.column.id, cell.row.index, value);
            };
         }

         input.value = String(cell.getValue());
         td.appendChild(input);
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

      let table = document.createElement("table");
      table.appendChild(this.#createThead());
      table.appendChild(this.#createTbody());

      this.appendChild(table);
   }

   /**
    * @param {string} columnId
    * @param {number} rowIndex
    * @param {string | number} value
    */
   #handleChange(columnId, rowIndex, value) {
      this.dispatchEvent(
         new CustomEvent("cellChange", {
            detail: {
               columnId,
               rowIndex,
               value
            },
            bubbles: true
         })
      );
   }

   /**
    * @param {EventListenerOrEventListenerObject} callback
    */
   onCellChange(callback) {
      this.addEventListener("cellChange", callback);
   }
}

customElements.define("data-table", DataTable);
