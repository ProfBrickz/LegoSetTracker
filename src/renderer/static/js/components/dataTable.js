// Imports
import { createTable, getCoreRowModel } from "@tanstack/table-core";
/** @import {ColumnDef,Table, Header, Row, Cell} from "@tanstack/table-core" */


// Types
/**
 * @typedef {Record<string, unknown>} TableRow
 */
/**
 * @typedef {ColumnDef<TableRow, any>} TableColumn
 */
/**
 * @typedef {Cell<TableRow, unknown>} TableCell
 */


// Web Component
export class DataTable extends HTMLElement {
   /** @type {Table<TableRow>} */
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
      this.table.setOptions(previous => ({
         ...previous,
         columns
      }));
   }

   /**
    * @param {TableRow[]} data
    */
   set data(data) {
      this.table.setOptions(previous => ({
         ...previous,
         data
      }));
   }

   /**
    * @private
    * @param {Header<TableRow, unknown>} header
    */
   createHeader(header) {
      let th = document.createElement("th");

      th.innerText = header.column.columnDef.header?.toString() || "";

      return th;
   }

   /**
    * @private
    */
   createThead() {
      let thead = document.createElement("thead");
      let row = document.createElement("tr");
      thead.appendChild(row);

      for (let header of this.table.getFlatHeaders()) {
         row.appendChild(this.createHeader(header));
      }

      return thead;
   }

   /**
    * @private
    * @param {TableCell} cell
    */
   createCell(cell) {
      let td = document.createElement("td");
      td.classList.add(cell.column.id);

      let { type, onChange, min, max, classList } = cell.column.columnDef.meta || {};

      if (typeof classList === "string") {
         td.classList.add(classList);
      } else if (typeof classList === "function") {
         td.classList.add(...classList({ row: cell.row }));
      }

      if (onChange && (type === "string" || type === "number")) {
         let input = document.createElement("input");

         if (type === "string") {
            input.type = "text";
            input.onchange = (event) => {
               let element = /** @type {HTMLInputElement} */ (event.target);
               if (!element) return;

               let value = element.value;
               this.data[cell.row.index][cell.column.id] = value;

               if (onChange) {
                  onChange({
                     element: /** @type {HTMLTableCellElement} */ (element.parentElement),
                     rowIndex: cell.row.index,
                     value,
                     row: cell.row
                  });
               }
            };
         } else if (type === "number") {
            input.type = "number";

            if (typeof min === "number") input.min = min.toString();
            else if (typeof min === "function") input.min = min(cell.row).toString();
            if (typeof max === "number") input.max = max.toString();
            else if (typeof max === "function") input.max = max(cell.row).toString();


            input.onchange = (event) => {
               let element = /** @type {HTMLInputElement} */ (event.target);
               if (!element) return;

               let value = element.valueAsNumber;
               this.data[cell.row.index][cell.column.id] = value;

               if (onChange) {
                  onChange({
                     element: /** @type {HTMLTableCellElement} */ (element.parentElement),
                     rowIndex: cell.row.index,
                     value,
                     row: cell.row
                  });
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
    * @private
    * @param {Row<TableRow>} row
   */
   createRow(row) {
      let tr = document.createElement("tr");

      for (let cell of row.getAllCells()) {
         tr.appendChild(this.createCell(cell));
      }

      return tr;
   }

   /**
    * @private
    */
   createTbody() {
      let tbody = document.createElement("tbody");

      for (let row of this.table.getRowModel().rows) {
         tbody.appendChild(this.createRow(row));
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
      tableElement.appendChild(this.createThead());
      tableElement.appendChild(this.createTbody());

      this.appendChild(tableElement);
   }

   /**
    * @param {number} rowIndex
    */
   deleteRow(rowIndex) {
      this.data = this.data.toSpliced(rowIndex, 1);

      this.renderTable();
   }
}

customElements.define("data-table", DataTable);
