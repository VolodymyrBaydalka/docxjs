import globalXmlParser from "../parser/xml-parser";
import { DocxElement } from "./dom";

export class BreakElement extends DocxElement {
    type: "page" | "lastRenderedPageBreak" | "textWrapping";

    protected parse(elem: Element) {
        super.parse(elem);
        
        if (elem.localName === "lastRenderedPageBreak") {
            this.type = "page";
        } else {
            this.type = <any>globalXmlParser.attr(elem, "type") ?? "textWrapping";
        }
    }
}