import { XmlParser } from "../parser/xml-parser";
import { Length } from "./common";
import { OpenXmlElement } from "./dom";
import { FooterPart } from "../footer/footer-part";

export interface Column {
    space: Length;
    width: Length;
}

export interface Columns {
    space: Length;
    numberOfColumns: number;
    separator: boolean;
    equalWidth: boolean;
    columns: Column[];
}

export interface PageSize {
    width: Length,
    height: Length,
    orientation: "landscape" | string
}

export interface PageMargins {
    top: Length;
    right: Length;
    bottom: Length;
    left: Length;
    header: Length;
    footer: Length;
    gutter: Length;
}

export enum SectionType {
    Continuous = "continuous",
    NextPage = "nextPage",
    NextColumn = "nextColumn",
    EvenPage = "evenPage",
    OddPage = "oddPage",
}

export interface FooterHeaderReference {
    id: string;
    type: string | "first" | "even" | "default";
}

export interface SectionProperties {
    type: SectionType | string;
    pageSize: PageSize,
    pageMargins: PageMargins,
    columns: Columns;
    id: string;
    footerRefs: FooterHeaderReference[];
    headerRefs: FooterHeaderReference[];
    forceFirstFooterHeaderDifferent: boolean;
}

export interface SectionRenderProperties extends SectionProperties {
    pageWithinSection: number;
}

export interface Section {
    elements: OpenXmlElement[],
    sectProps: SectionProperties
}

export function parseSectionProperties(elem: Element, xml: XmlParser): SectionProperties {
    const section = <SectionProperties>{footerRefs: [], headerRefs: []};
    section.id = xml.attr(elem, "rsidSect");

    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "pgSz":
                section.pageSize = {
                    width: xml.lengthAttr(e, "w"),
                    height: xml.lengthAttr(e, "h"),
                    orientation: xml.attr(e, "orient")
                }
                break;

            case "type":
                section.type = xml.attr(e, "val");
                break;

            case "pgMar":
                section.pageMargins = {
                    left: xml.lengthAttr(e, "left"),
                    right: xml.lengthAttr(e, "right"),
                    top: xml.lengthAttr(e, "top"),
                    bottom: xml.lengthAttr(e, "bottom"),
                    header: xml.lengthAttr(e, "header"),
                    footer: xml.lengthAttr(e, "footer"),
                    gutter: xml.lengthAttr(e, "gutter"),
                };
                break;

            case "cols":
                section.columns = parseColumns(e, xml);
                break;

            case "titlePg":
                const titlePageVal: string = xml.attr(e, "val");
                if(titlePageVal !== "false") {
                    section.forceFirstFooterHeaderDifferent = true;
                }
                break;

            case "headerReference":
                section.headerRefs.push(parseFooterHeaderReference(e, xml)); 
                break;
            
            case "footerReference":
                section.footerRefs.push(parseFooterHeaderReference(e, xml)); 
                break;
        }
    }

    return section;
}

function parseColumns(elem: Element, xml: XmlParser): Columns {
    return {
        numberOfColumns: xml.intAttr(elem, "num"),
        space: xml.lengthAttr(elem, "space"),
        separator: xml.boolAttr(elem, "sep"),
        equalWidth: xml.boolAttr(elem, "equalWidth", true),
        columns: xml.elements(elem, "col")
            .map(e => <Column>{
                width: xml.lengthAttr(e, "w"),
                space: xml.lengthAttr(e, "space")
            })
    };
}

function parseFooterHeaderReference(elem: Element, xml: XmlParser): FooterHeaderReference {
    return {
        id: xml.attr(elem, "id"),
        type: xml.attr(elem, "type"),
    }
}