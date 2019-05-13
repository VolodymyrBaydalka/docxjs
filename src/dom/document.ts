import { Length, Columns } from "./common";
import { OpenXmlElement } from "./dom";

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

export interface SectionProperties {
    pageSize: PageSize,
    pageMargins: PageMargins,
    columns: Columns;
}

export interface WordDocument extends OpenXmlElement {
    section: SectionProperties;
}