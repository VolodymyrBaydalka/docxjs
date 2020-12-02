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

export enum SectionType {
    Continuous = "continuous",
    NextPage = "nextPage", 
    NextColumn = "nextColumn",
    EvenPage = "evenPage",
    OddPage = "oddPage",
}

export interface SectionProperties {
    type: SectionType | string;
    pageSize: PageSize,
    pageMargins: PageMargins,
    columns: Columns;
}

export interface DocumentElement extends OpenXmlElement {
    props: SectionProperties;
}