import { DomType, OpenXmlElement } from "../document/dom";

export abstract class WmlBaseNote implements OpenXmlElement {
    id: string;
	type: DomType;
	noteType: string;
    children?: OpenXmlElement[] = [];
    cssStyle?: Record<string, string> = {};
    className?: string;
    parent?: OpenXmlElement;
}

export class WmlFootnote extends WmlBaseNote {
	type = DomType.Footnote
}

export class WmlEndnote extends WmlBaseNote {
	type = DomType.Endnote
}