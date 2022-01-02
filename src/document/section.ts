import globalXmlParser, { XmlParser } from "../parser/xml-parser";
import { Borders, parseBorders } from "./border";
import { Length } from "./common";

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

export interface PageNumber {
    start: number;
    chapSep: "colon" | "emDash" | "endash" | "hyphen" | "period" | string;
    chapStyle: string;
    format: "none" | "cardinalText" | "decimal" | "decimalEnclosedCircle" | "decimalEnclosedFullstop" 
        | "decimalEnclosedParen" | "decimalZero" | "lowerLetter" | "lowerRoman"
        | "ordinalText" | "upperLetter" | "upperRoman" | string;
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
    pageBorders: Borders;
    pageNumber: PageNumber;
    columns: Columns;
    footerRefs: FooterHeaderReference[];
    headerRefs: FooterHeaderReference[];
    titlePage: boolean;
}

export function parseSectionProperties(elem: Element, xml: XmlParser = globalXmlParser): SectionProperties {
    var section = <SectionProperties>{};

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

            case "headerReference":
                (section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml)); 
                break;
            
            case "footerReference":
                (section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml)); 
                break;

            case "titlePg":
                section.titlePage = xml.boolAttr(e, "val", true);
                break;

            case "pgBorders":
                section.pageBorders = parseBorders(e, xml);
                break;

            case "pgNumType":
                section.pageNumber = parsePageNumber(e, xml);
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

function parsePageNumber(elem: Element, xml: XmlParser): PageNumber {
    return {
        chapSep: xml.attr(elem, "chapSep"),
        chapStyle: xml.attr(elem, "chapStyle"),
        format: xml.attr(elem, "fmt"),
        start: xml.intAttr(elem, "start")
    };
}

function parseFooterHeaderReference(elem: Element, xml: XmlParser): FooterHeaderReference {
    return {
        id: xml.attr(elem, "id"),
        type: xml.attr(elem, "type"),
    }
}