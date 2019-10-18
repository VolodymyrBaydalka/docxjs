import { OpenXmlElement } from "./dom";
import { CommonProperties, Length, Borders } from "./common";
import { SectionProperties } from "./document";

export interface ParagraphElement extends OpenXmlElement {
    props: ParagraphProperties;
}

export interface ParagraphProperties extends CommonProperties {
    sectionProps: SectionProperties;
    tabs: ParagraphTab[];
    numbering: ParagraphNumbering;

    border: Borders;
    lineSpacing: LineSpacing;
}

export interface ParagraphTab {
    style: "bar" | "center" | "clear" | "decimal" | "end" | "num" | "start" | "left" | "right";
    leader: "none" | "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
    position: Length;
}

export interface ParagraphNumbering {
    id: string;
    level: number;
}

export interface LineSpacing {
    after: Length;
    before: Length;
    line: number;
    lineRule: "atLeast" | "exactly" | "auto";
}