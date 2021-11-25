import { DomType, OpenXmlElement } from "../document/dom";

export class WmlFootnote implements OpenXmlElement {
    id: string;
    footnoteType: string;
    type: DomType = DomType.Footnote;
    children?: OpenXmlElement[] = [];
    cssStyle?: Record<string, string> = {};
    className?: string;
    parent?: OpenXmlElement;
}