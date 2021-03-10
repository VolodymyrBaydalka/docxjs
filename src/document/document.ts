import { element, fromElement } from "../parser/xml-serialize";
import { DocxContainer, DocxElement } from "./dom";
import { parseSectionProperties, SectionProperties } from "./section";

@element("document")
export class DocumentElement extends DocxElement {
    body: BodyElement;
}

@element("body")
export class BodyElement extends DocxContainer {
    @fromElement("sectPr", parseSectionProperties)
    sectionProps: SectionProperties;
}