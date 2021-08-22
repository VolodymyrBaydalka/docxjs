import globalXmlParser, { XmlParser } from "../parser/xml-parser";
import { DocxContainer } from "./dom";

export class WmlDrawing extends DocxContainer {

}

export class DmlPicture extends DocxContainer {
    resourceId: string;
    stretch: any;
    offset: any;
    size: any;
}

export function parseDmlPicture(elem: Element, output: DmlPicture, xml: XmlParser = globalXmlParser) {
    const blipFill = xml.element(elem, "blipFill");
    const blip = xml.element(blipFill, "blip");

    output.resourceId = xml.attr(blip, "embed");
}