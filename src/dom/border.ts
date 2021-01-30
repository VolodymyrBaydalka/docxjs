import { XmlParser } from "../parser/xml-parser";
import { Length, LengthUsage } from "./common";

export interface Border {
    color: string;
    type: string;
    size: Length;
    frame: boolean;
    shadow: boolean;
    offset: Length;
}

export interface Borders {
    top: Border;
    left: Border;
    right: Border;
    bottom: Border;
}

export function parseBorder(elem: Element, xml: XmlParser): Border {
    return {
        type: xml.attr(elem, "val"),
        color: xml.attr(elem, "color"),
        size: xml.lengthAttr(elem, "sz", LengthUsage.Border),
        offset: xml.lengthAttr(elem, "space", LengthUsage.Point),
        frame: xml.boolAttr(elem, 'frame'),
        shadow: xml.boolAttr(elem, 'shadow')
    };
}

export function parseBorders(elem: Element, xml: XmlParser): Borders {
    var result = <Borders>{};

    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "left": result.left = parseBorder(e, xml); break;
            case "top": result.top = parseBorder(e, xml); break;
            case "right": result.right = parseBorder(e, xml); break;
            case "bottom": result.bottom = parseBorder(e, xml); break;
        }
    }

    return result;
}