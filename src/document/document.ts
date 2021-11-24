import { OpenXmlElement } from "./dom";
import { SectionProperties } from "./section";

export interface DocumentElement extends OpenXmlElement {
    props: SectionProperties;
}