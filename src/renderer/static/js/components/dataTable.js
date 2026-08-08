// Imports
import { createTable, functionalUpdate, getCoreRowModel, getExpandedRowModel } from "@tanstack/table-core";
/** @import {ColumnDef,Table, Header, Row, Cell} from "@tanstack/table-core" */


// Types
/** @typedef {{rowId: number | undefined, subRows: subRow[]}} subRow  */
/** @typedef {Record<string, unknown> & {rowId?: number, subRows: subRow[], grayedOut?: boolean}} TableRow */
/** @typedef {ColumnDef<TableRow, any>} TableColumn */
/** @typedef {Cell<TableRow, unknown>} TableCell */


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
         state: {
            columnPinning: {},
            expanded: {}
         },
         getSubRows: (row) => /** @type {TableRow[]} */(row.subRows),
         getCoreRowModel: getCoreRowModel(),
         getExpandedRowModel: getExpandedRowModel(),
         onStateChange:
            /**
             * @param {any} updater
             */
            (updater) => {
               const newState = functionalUpdate(updater, this.table.options.state);
               this.table.options.state = newState;
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
    * @param {subRow[]} subRows
    * @param {number} parentRowId
    * @param {boolean} isExpanded
    */
   expandCollapseSubRows(subRows, parentRowId, isExpanded) {
      for (let subRow of subRows) {
         let rowId = subRow.rowId;
         let tableRow =
         /** @type {HTMLTableRowElement} */ (document.querySelector(`[data-row-id="${rowId}"][data-parent-row-id="${parentRowId}"]`));

         if (isExpanded) {
            tableRow.style.display = "";
         } else {
            tableRow.style.display = "none";

            let expandButton = /** @type {HTMLButtonElement | undefined} */ (tableRow.querySelector("td.expand button"));
            if (!!expandButton) expandButton.classList.remove("expanded");

            if (subRow.rowId == undefined) return;
            if (subRow.subRows.length > 0) this.expandCollapseSubRows(subRow.subRows, subRow.rowId, isExpanded);
         }
      }
   }

   /**
    * @param {TableColumn[]} columns
    */
   set columns(columns) {
      columns.unshift({
         id: "expand",
         header: "",
         cell: ({ row }) => {
            let fragment = document.createDocumentFragment();

            if (!row.getCanExpand()) return fragment;

            /** @type {HTMLDivElement | null} */
            let div = null;
            if (row.depth > 0) div = document.createElement("div");

            let expandButton = document.createElement("button");

            let collapsedIcon = document.createElement("i");
            collapsedIcon.classList.add("collapsed-icon");
            collapsedIcon.dataset.lucide = "chevron-down";
            expandButton.appendChild(collapsedIcon);

            expandButton.onclick = async (event) => {
               event.stopPropagation();
               await row.getToggleExpandedHandler()();
               let isExpanded = row.getIsExpanded();

               if (row.original.rowId == undefined) return;
               this.expandCollapseSubRows(row.original.subRows, row.original.rowId, isExpanded);

               expandButton.classList.toggle("expanded");
            };

            if (div !== null) {
               div.appendChild(expandButton);
               fragment.appendChild(div);
            } else {
               fragment.appendChild(expandButton);
            }

            return fragment;
         },
         meta: {
            type: "function"
         }
      });

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

      let rowId = row.original.rowId;
      if (rowId !== undefined) {
         tr.dataset.rowId = rowId.toString();
      }

      if (row.original.grayedOut) {
         tr.classList.add("grayed-out");
      }

      tr.dataset.depth = row.depth.toString();

      for (let cell of row.getAllCells()) {
         tr.appendChild(this.createCell(cell));
      }

      return tr;
   }

   /**
    * @param {HTMLTableSectionElement} tbody
    * @param {Row<TableRow>} row
    */
   createSubRows(tbody, row) {
      if (row.subRows.length <= 0) return;
      if (row.original.rowId == undefined) return;

      for (let i = 0; i < row.subRows.length; i++) {
         let subRow = row.subRows[i];

         let htmlSubRow = this.createRow(subRow);
         htmlSubRow.style.display = "none";
         htmlSubRow.dataset.parentRowId = row.original.rowId.toString();

         if (i == row.subRows.length - 1) {
            htmlSubRow.classList.add("last-sub-row");
         }

         tbody.appendChild(htmlSubRow);

         this.createSubRows(tbody, subRow);
      }
   }

   /**
    * @private
    */
   createTbody() {
      let tbody = document.createElement("tbody");

      for (let row of this.table.getRowModel().rows) {
         let htmlRow = this.createRow(row);

         tbody.appendChild(htmlRow);

         this.createSubRows(tbody, row);
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
      // tableElement.classList.add("scroll-area");

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
