import { OpenXmlElement } from "./dom";
import { CommonProperties, Length } from "./common";
import { SectionProperties } from "./document";

export interface ParagraphElement extends OpenXmlElement {
    props: ParagraphProperties;
}

export interface ParagraphProperties extends CommonProperties {
    sectionProps: SectionProperties;
    tabs: ParagraphTab[];
    numbering: ParagraphNumbering;
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