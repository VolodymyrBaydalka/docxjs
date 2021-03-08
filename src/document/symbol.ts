import { element, fromAttribute } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element('sym')
export class SymbolElement extends DocxElement {
    @fromAttribute('font')
    font: string;
    @fromAttribute('char')
    char: string;
}