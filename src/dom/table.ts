import { DocxContainer } from "./dom";

export class TableElement extends DocxContainer {
    columns?: TableColumn[];
    cellStyle?: Record<string, string>;
}

export interface TableColumn {
    width?: string;
}