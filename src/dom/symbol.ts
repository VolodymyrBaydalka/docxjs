import globalXmlParser from "../parser/xml-parser";
import { DocxElement } from "./dom";

export class SymbolElement extends DocxElement {
    font: string;
    char: string;

    protected parse(elem: Element) {
        super.parse(elem);
        this.font = globalXmlParser.attr(elem, "font");
        this.char = globalXmlParser.attr(elem, "char");
    }
}