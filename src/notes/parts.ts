import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { WmlBaseNote, WmlEndnote, WmlFootnote } from "./elements";

export class BaseNotePart<T extends WmlBaseNote> extends Part {
    protected _documentParser: DocumentParser;

    notes: T[]

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }
}

export class FootnotesPart extends BaseNotePart<WmlFootnote> {
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path, parser);
    }

    parseXml(root: Element) {
        this.notes = this._documentParser.parseNotes(root, "footnote", WmlFootnote);
    }
}

export class EndnotesPart extends BaseNotePart<WmlEndnote> {
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path, parser);
    }

    parseXml(root: Element) {
        this.notes = this._documentParser.parseNotes(root, "endnote", WmlEndnote);
    }
}