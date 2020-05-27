import { ParagraphTab, ParagraphProperties, ParagraphNumbering, LineSpacing } from "../dom/paragraph";
import * as xml from "./common";
import { ns } from "../dom/common";
import { parseSectionProperties } from "./section";
import { convertLength } from "./common";
import { buildXmlSchema, deserializeSchema } from "./xml-serialize";

export function parseParagraphProperties(elem: Element, props: ParagraphProperties) {
    if (elem.namespaceURI != ns.wordml)
        return false;

    switch (elem.localName) {
        case "tabs":
            props.tabs = parseTabs(elem);
            break;

        case "sectPr":
            props.sectionProps = parseSectionProperties(elem);
            break;

        case "numPr":
            props.numbering = deserializeSchema(elem, {}, numberingSchema);
            break;
        
        case "spacing":
            props.lineSpacing = deserializeSchema(elem, {}, lineSpacingSchema);
            return false; // TODO
            break;

        default:
            return false;
    }

    return true;
}

function parseTabs(elem: Element): ParagraphTab[] {
    return xml.elements(elem, ns.wordml, "tab")
        .map(e => <ParagraphTab>{
            position: xml.lengthAttr(e, ns.wordml, "pos"),
            leader: xml.stringAttr(e, ns.wordml, "leader"),
            style: xml.stringAttr(e, ns.wordml, "val")
        });
}

const numberingSchema = buildXmlSchema({
    $elem: "numPr",
    id: { $attr: "numId" },
    level: { $attr: "ilvl", convert: (v) => parseInt(v) },
})

const lineSpacingSchema = buildXmlSchema({
    $elem: "spacing",
    before: { $attr: "before", convert: (v) => convertLength(v) },
    after: { $attr: "after", convert: (v) => convertLength(v) },
    line: { $attr: "line", convert: (v) => parseInt(v) },
    lineRule: { $attr: "before" },
});