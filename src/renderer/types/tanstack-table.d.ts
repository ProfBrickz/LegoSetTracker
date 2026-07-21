import "@tanstack/table-core";
import { Row } from "@tanstack/table-core";
import { TableRow } from "../static/js/components/dataTable.js";

declare module "@tanstack/table-core" {
	export interface ColumnMeta<TData extends RowData, TValue> {
		type: "string" | "number" | "image" | "function";
		min?: number | ((row: Row<TableRow>) => number);
		max?: number | ((row: Row<TableRow>) => number);
		onChange?: (options: { element: HTMLTableCellElement, rowIndex: number, value: string | number, row: Row<TableRow>; }) => void;
		classList?: string[] | ((options: { row: Row<TableRow>; }) => string[]);
	}
}
