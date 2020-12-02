import { XmlParser } from "../parser/xml-parser";

export interface Relationship {
    id: string,
    type: RelationshipTypes | string,
    target: string
    targetMode: "" | string 
}

export enum RelationshipTypes {
    OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    FontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    Image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    Styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    StylesWithEffects = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
    Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    WebSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
    Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
}

export function parseRelationships(root: Element, xmlParser: XmlParser): Relationship[] {
    return xmlParser.elements(root).map(e => <Relationship>{
        id: xmlParser.attr(e, "Id"),
        type: xmlParser.attr(e, "Type"),
        target: xmlParser.attr(e, "Target"),
        targetMode: xmlParser.attr(e, "TargetMode")
    });
}