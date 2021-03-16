import { element, fromElement } from "../parser/xml-serialize";
import { DocxContainer, DocxElement } from "./dom";
import { parseSectionProperties, SectionProperties } from "./section";

@element("document")
export class WmlDocument extends DocxElement {
    body: WmlBody;
}

@element("body")
export class WmlBody extends DocxContainer {
    @fromElement("sectPr", parseSectionProperties)
    sectionProps: SectionProperties;
}