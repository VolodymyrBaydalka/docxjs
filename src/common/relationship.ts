import { XmlParser } from "../parser/xml-parser";

export interface Relationship {
    id: string,
    type: RelationshipTypes | string,
    target: string
    targetMode: "" | "External" | string 
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
    Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    Footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
	Endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    Footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    Header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
	CustomProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties",
	Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    CommentsExtended = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    AltChunk = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
}

export function parseRelationships(root: Element, xml: XmlParser): Relationship[] {
    return xml.elements(root).map(e => <Relationship>{
        id: xml.attr(e, "Id"),
        type: xml.attr(e, "Type"),
        target: xml.attr(e, "Target"),
        targetMode: xml.attr(e, "TargetMode")
    });
}