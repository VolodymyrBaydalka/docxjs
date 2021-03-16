import { element } from "../parser/xml-serialize";
import { DocxContainer } from "./dom";

@element("tr")
export class WmlTableRow extends DocxContainer {
}