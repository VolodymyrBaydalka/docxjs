import { OpenXmlElement, DocxTab } from "./dom";
import { CommonProperties } from "./common";
import { SectionProperties } from "./document";

export interface ParagraphElement extends OpenXmlElement {
    numberingId?: string;
    numberingLevel?: number;
    tabs: DocxTab[];
    
    props: ParagraphProperties;
}

export interface ParagraphProperties extends CommonProperties {
    sectionProps: SectionProperties;
}