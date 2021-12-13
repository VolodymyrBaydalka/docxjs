import { XmlParser } from "../parser/xml-parser";
import { Length } from "../document/common";

export class DmlSettings {
    defaultTabStopWidth: Length;
}



export function parseSettings(elem: Element, xml: XmlParser) {
    var result = new DmlSettings();
    const defaultTabStopElement = xml.element(elem, "defaultTabStop");
    if(defaultTabStopElement) {
        result.defaultTabStopWidth = xml.lengthAttr(defaultTabStopElement, "val");
    }


    return result;
}