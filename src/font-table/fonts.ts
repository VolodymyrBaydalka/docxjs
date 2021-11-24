import { XmlParser } from "../parser/xml-parser";

export interface FontDeclaration {
    name: string,
    altName: string,
    family: string,
    fontKey: string,
    refId: string
}

export function parseFonts(root: Element, xmlParser: XmlParser): FontDeclaration[] {
    return xmlParser.elements(root).map(el => parseFont(el, xmlParser));
}

export function parseFont(elem: Element, xmlParser: XmlParser): FontDeclaration {
    let result = <FontDeclaration>{
        name: xmlParser.attr(elem, "name")
    };

    for (let el of xmlParser.elements(elem)) {
        switch (el.localName) {
            case "family":
                result.family = xmlParser.attr(el, "val");
                break;

            case "altName":
                result.altName = xmlParser.attr(el, "val");
                break;

            case "embedRegular":
                result.fontKey = xmlParser.attr(el, "fontKey");
                result.refId = xmlParser.attr(el, "id");
                break;
        }
    }

    return result;
}