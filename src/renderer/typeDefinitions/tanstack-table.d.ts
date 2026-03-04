import "@tanstack/table-core";

declare module "@tanstack/table-core" {
	export interface ColumnMeta<TData extends RowData, TValue> {
		editable?: boolean;
		type: "string" | "number" | "image";
		onChange?: (value: string | number) => void;
	}
}
