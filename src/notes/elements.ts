import { OpenXmlElementBase, DomType } from "../document/dom";

export abstract class WmlBaseNote implements OpenXmlElementBase {
    type: DomType;
    id: string;
	noteType: string;
}

export class WmlFootnote extends WmlBaseNote {
	type = DomType.Footnote
}

export class WmlEndnote extends WmlBaseNote {
	type = DomType.Endnote
}