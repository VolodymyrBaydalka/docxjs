import { element, fromAttribute } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element('br')
export class WmlBreak extends DocxElement {
    @fromAttribute("type")
    type: "page" | "column" | "textWrapping" = "textWrapping";
    
    @fromAttribute("clear")
    clear: "all" | "left" | "right" | "none";
}

@element('lastRenderedPageBreak')
export class WmlLastRenderedPageBreak extends DocxElement {
}