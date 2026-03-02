import { createTable, getCoreRowModel, } from "@tanstack/table-core";

class DataTable extends HTMLElement {
   /** @type {boolean} */
   initialized = false;

   constructor() {
      super();
      this.innerHTML = "hi";

      /** @type {import("@tanstack/table-core").ColumnDef<Record<string, unknown>, any>[]} */
      let columns = JSON.parse(this.getAttribute("columns") || "[]");
      /** @type {Record<string, unknown>[]} */
      let data = JSON.parse(this.getAttribute("data") || "[]");

      this.table = createTable({
         columns,
         data,
         getCoreRowModel: getCoreRowModel(),
         onStateChange: () => { },
         state: {
            columnPinning: {}
         },
         renderFallbackValue: null
      });

      this.#createTable();
   }

   static observedAttributes = ["data"];

   /**
    * @param {string} name
    * @param {string} newValue
    * @param {string} oldValue
    */
   attributeChangedCallback(name, oldValue, newValue) {
      if (!this.initialized) return;
      if (name !== "data") return;

      let data = JSON.parse(newValue);

      this.#setData(data);
   }

   connectedCallback() {
      this.initialized = true;
   }

   /**
    * @param {Record<string, unknown>[]} data
    */
   #setData(data) {
      this.table.setOptions(previous => {
         console.log("oldData", previous.data);
         console.log("newData", data);

         return ({
            ...previous,
            data
         });
      });
   }

   /**
    *
    * @param {import("@tanstack/table-core").Header<Record<string, unknown>, unknown>} header
    */
   #createHeader(header) {
      let th = document.createElement("th");

      th.innerText = header.id;

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
    * @param {import("@tanstack/table-core").Row<Record<string, unknown>>} row
   */
   #createRow(row) {
      let tr = document.createElement("tr");

      for (let cell of row.getAllCells()) {
         let td = document.createElement("td");
         tr.appendChild(td);

         td.innerText = String(cell.getValue());
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

   #createTable() {
      this.innerHTML = "";

      let table = document.createElement("table");
      table.appendChild(this.#createThead());
      table.appendChild(this.#createTbody());

      this.appendChild(table);
   }
}

customElements.define("data-table", DataTable);
