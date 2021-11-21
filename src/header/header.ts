import { DomType, OpenXmlElement } from "../dom/dom";

export class WmlHeader implements OpenXmlElement {
    type: DomType = DomType.Header;
    children?: OpenXmlElement[] = [];
    cssStyle?: Record<string, string> = {};
    className?: string;
    parent?: OpenXmlElement;
}