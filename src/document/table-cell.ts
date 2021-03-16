import { element } from "../parser/xml-serialize";
import { DocxContainer } from "./dom";

@element("tc")
export class WmlTableCell extends DocxContainer {
    span?: number;
    verticalMerge: string;
}