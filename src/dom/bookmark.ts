import { XmlParser } from "../parser/xml-parser";
import { DomType, OpenXmlElement } from "./dom";

export interface BookmarkStartElement extends OpenXmlElement {
    id: string;
    name: string;
    colFirst: number;
    colLast: number;
}

export interface BookmarkEndElement extends OpenXmlElement {
    id: string;
}

export function parseBookmarkStart(elem: Element, xml: XmlParser): BookmarkStartElement {
    return {
        type: DomType.BookmarkStart,
        id: xml.attr(elem, "id"),
        name: xml.attr(elem, "name"),
        colFirst: xml.intAttr(elem, "colFirst"),
        colLast: xml.intAttr(elem, "colLast")
    }
}

export function parseBookmarkEnd(elem: Element, xml: XmlParser): BookmarkEndElement {
    return {
        type: DomType.BookmarkEnd,
        id: xml.attr(elem, "id")
    }
}