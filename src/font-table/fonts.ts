import { XmlParser } from "../parser/xml-parser";

const embedFontTypeMap = {
    embedRegular: 'regular',
    embedBold: 'bold',
    embedItalic: 'italic',
    embedBoldItalic: 'boldItalic',
}

export interface FontDeclaration {
    name: string,
    altName: string,
    family: string,
    embedFontRefs: EmbedFontRef[];
}

export interface EmbedFontRef {
    id: string;
    key: string;
    type: 'regular' | 'bold' | 'italic' | 'boldItalic';
}

export function parseFonts(root: Element, xml: XmlParser): FontDeclaration[] {
    return xml.elements(root).map(el => parseFont(el, xml));
}

export function parseFont(elem: Element, xml: XmlParser): FontDeclaration {
    let result = <FontDeclaration>{
        name: xml.attr(elem, "name"),
        embedFontRefs: []
    };

    for (let el of xml.elements(elem)) {
        switch (el.localName) {
            case "family":
                result.family = xml.attr(el, "val");
                break;

            case "altName":
                result.altName = xml.attr(el, "val");
                break;

            case "embedRegular":
            case "embedBold":
            case "embedItalic":
            case "embedBoldItalic":
                result.embedFontRefs.push(parseEmbedFontRef(el, xml));
                break;
        }
    }

    return result;
}

export function parseEmbedFontRef(elem: Element, xml: XmlParser): EmbedFontRef {
    return { 
        id: xml.attr(elem, "id"), 
        key: xml.attr(elem, "fontKey"),
        type: embedFontTypeMap[elem.localName]
    };
}