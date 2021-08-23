import { XmlParser } from "../parser/xml-parser";
import { Length } from "./common";

export interface Indentation {
    start: Length;
    end: Length;
    hanging: Length;
    firstLine: Length;
}

export function parseIndentation(elem: Element, xml: XmlParser): Indentation {
    return {
        start: xml.lengthAttr(elem, "start"),
        end: xml.lengthAttr(elem, "end"),
        hanging: xml.lengthAttr(elem, "hanging"),
        firstLine: xml.lengthAttr(elem, "firstLine"),
    } as Indentation;
}