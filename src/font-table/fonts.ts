import { XmlParser } from "../parser/xml-parser";

export interface FontDeclatation {
    name: string,
    fontKey?: string,
    refId?: string
}

export function parseFonts(root: Element, xmlParser: XmlParser): FontDeclatation[] {
    const result = [];

    for(let el of xmlParser.elements(root)) {
        let font: FontDeclatation = {
            name: xmlParser.attr(el, "name")
        }

        let embed = xmlParser.element(el, "embedRegular");

        if(embed) {
            font.fontKey = xmlParser.attr(embed, "fontKey");    
            font.refId = xmlParser.attr(embed, "id");    
        }

        result.push(font);
    }

    return result;
}