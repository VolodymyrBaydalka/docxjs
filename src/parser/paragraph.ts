import { ParagraphTab, ParagraphProperties, ParagraphNumbering, LineSpacing } from "../dom/paragraph";
import * as xml from "./common";
import { ns } from "../dom/common";
import { parseSectionProperties } from "./section";

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
            props.numbering = parseNumbering(elem);
            break;
        
        case "spacing":
            props.lineSpacing = parseLineSpacing(elem);
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

function parseNumbering(elem: Element): ParagraphNumbering {
    var result = <ParagraphNumbering>{};

    for (let e of xml.elements(elem, ns.wordml)) {
        switch (e.localName) {
            case "numId":
                result.id = xml.stringAttr(e, ns.wordml, "val");
                break;

            case "ilvl":
                result.level = xml.intAttr(e, ns.wordml, "val");
                break;
        }
    }

    return result;
}

function parseLineSpacing(elem: Element): LineSpacing {
    return {
        before: xml.lengthAttr(elem, ns.wordml, "before"),
        after: xml.lengthAttr(elem, ns.wordml, "after"),
        line: xml.intAttr(elem, ns.wordml, "line"),
        lineRule: xml.stringAttr(elem, ns.wordml, "lineRule")
    } as LineSpacing;
}