import { element } from "../parser/xml-serialize";
import { DocxContainer } from "./dom";

@element("table")
export class TableElement extends DocxContainer {
    columns?: TableColumn[];
    cellStyle?: Record<string, string>;
}

export interface TableColumn {
    width?: string;
}