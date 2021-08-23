import { XmlParser } from "../parser/xml-parser";
import { Length, LengthUsage } from "./common";

export interface LineSpacing {
    after: Length;
    before: Length;
    line: Length;
    lineRule: "atLeast" | "exactly" | "auto";
}

export function parseLineSpacing(elem: Element, xml: XmlParser): LineSpacing {
    const lineRule = xml.attr(elem, "lineRule");

    return {
        before: xml.lengthAttr(elem, "before"),
        after: xml.lengthAttr(elem, "after"),
        line: xml.lengthAttr(elem, "line", lineRule === 'auto' ? LengthUsage.LineHeight : LengthUsage.Dxa),
        lineRule
    } as LineSpacing;
}