import { element, fromAttribute, fromText } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element('t')
export class WmlText extends DocxElement {
    @fromText()
    text: string;
}

@element('sym')
export class WmlSymbol extends DocxElement {
    @fromAttribute('font')
    font: string;
    @fromAttribute('char')
    char: string;
}

@element('tab')
export class WmlTab extends DocxElement {
}

@element("instrText")
export class WmlInstructionText extends DocxElement {
    @fromText()
    text: string;
}