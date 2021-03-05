import { element } from "../parser/xml-serialize";
import { DocxContainer } from "./dom";

@element("tc")
export class TableCellElement extends DocxContainer {
    span?: number;
}