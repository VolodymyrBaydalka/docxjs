import { XmlParser } from "../parser/xml-parser";
import { Length } from "./common";

export interface LineSpacing {
    after: Length;
    before: Length;
    line: number;
    lineRule: "atLeast" | "exactly" | "auto";
}

export function parseLineSpacing(elem: Element, xml: XmlParser): LineSpacing {
    return {
        before: xml.lengthAttr(elem, "before"),
        after: xml.lengthAttr(elem, "after"),
        line: xml.intAttr(elem, "line"),
        lineRule: xml.attr(elem, "lineRule")
    } as LineSpacing;
}