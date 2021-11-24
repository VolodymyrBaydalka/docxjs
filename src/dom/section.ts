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

export interface SectionFooter {
    forceFirstDifferent?: boolean;
    default?: string;
    first?: string;
    even?: string;
}

export interface SectionProperties {
    type: SectionType | string;
    pageSize: PageSize,
    pageMargins: PageMargins,
    columns: Columns;
    footer: SectionFooter;
    id: string;
}

export interface SectionRenderProperties extends SectionProperties {
    pageWithinSection: number;
}

export interface Section {
    elements: OpenXmlElement[],
    sectProps: SectionProperties
}

export function parseSectionProperties(elem: Element, xml: XmlParser): SectionProperties {
    const section = <SectionProperties>{footer: {}};
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

            case "footerReference":
                const footerType: string = xml.attr(e, "type");
                const footerId: string = xml.attr(e, "id");
                switch (footerType) {
                    case "default":
                        section.footer.default = footerId;
                        break;
                    case "first":
                        section.footer.first = footerId;
                        break;
                    case "even":
                        section.footer.even = footerId;
                        break;
                }
                break;
            case "titlePg":
                const titlePageVal: string = xml.attr(e, "val");
                if(titlePageVal !== "false") {
                    section.footer.forceFirstDifferent = true;
                }
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