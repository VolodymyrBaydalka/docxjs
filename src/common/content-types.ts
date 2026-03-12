import { XmlParser } from "../parser/xml-parser";

export interface ContentType {
    extension: string,
    partName: string,
    contentType: string;
}

export function parseContentTypes(root: Element, xml: XmlParser): ContentType[] {
    return xml.elements(root).map(e => <ContentType>{
        extension: xml.attr(e, "Extension"),
        partName: xml.attr(e, "PartName"),
        contentType: xml.attr(e, "ContentType")
    });
}