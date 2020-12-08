import { XmlParser } from "../parser/xml-parser";
import { Length, LengthUsage } from "./common";

export interface Border {
    color: string;
    type: string;
    size: Length;
}

export interface Borders {
    top: Border;
    left: Border;
    right: Border;
    botton: Border;
}

export function parseBorder(elem: Element, xml: XmlParser): Border {
    return {
        type: xml.attr(elem, "val"),
        color: xml.attr(elem, "color"),
        size: xml.lengthAttr(elem, "sz", LengthUsage.Border)
    };
}

export function parseBorders(elem: Element, xml: XmlParser): Borders {
    var result = <Borders>{};

    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "left": result.left = parseBorder(e, xml); break;
            case "top": result.top = parseBorder(e, xml); break;
            case "right": result.right = parseBorder(e, xml); break;
            case "botton": result.botton = parseBorder(e, xml); break;
        }
    }

    return result;
}