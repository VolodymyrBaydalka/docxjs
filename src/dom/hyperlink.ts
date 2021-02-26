import globalXmlParser from "../parser/xml-parser";
import { DocxContainer } from "./dom";

export class HyperlinkElement extends DocxContainer {
    href?: string;

    protected parse(elem: Element) {
        this.href = globalXmlParser.attr(elem, "anchor");
    }
}