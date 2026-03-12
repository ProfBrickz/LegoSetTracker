import "@tanstack/table-core";

declare module "@tanstack/table-core" {
	export interface ColumnMeta<TData extends RowData, TValue> {
		type: "string" | "number" | "image" | "function";
		editable?: boolean;
		min?: number;
		max?: number;
		onChange?: (rowIndex: number, value: string | number) => void;
	}
}
