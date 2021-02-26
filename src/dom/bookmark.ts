import globalXmlParser, { XmlParser } from "../parser/xml-parser";
import { DocxElement } from "./dom";

export class BookmarkStartElement extends DocxElement {
    id: string;
    name: string;
    colFirst: number;
    colLast: number;

    protected parse(elem: Element) {
        super.parse(elem);
        this.id = globalXmlParser.attr(elem, "id");
        this.name = globalXmlParser.attr(elem, "name");
        this.colFirst = globalXmlParser.intAttr(elem, "colFirst"),
        this.colLast = globalXmlParser.intAttr(elem, "colLast")
    }
}

export class BookmarkEndElement extends DocxElement {
    id: string;

    protected parse(elem: Element) {
        super.parse(elem);
        this.id = globalXmlParser.attr(elem, "id");
    }
}