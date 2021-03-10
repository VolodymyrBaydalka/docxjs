import { element, fromAttribute } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element('br')
export class BreakElement extends DocxElement {
    @fromAttribute("type")
    type: "page" | "column" | "textWrapping";
    
    @fromAttribute("clear")
    clear: "all" | "left" | "right" | "none";
}

@element('lastRenderedPageBreak')
export class LastRenderedPageBreakElement extends DocxElement {
}