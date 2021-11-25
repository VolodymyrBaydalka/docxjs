import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { WmlFootnote } from "./footnote";

export class FootnotesPart extends Part {
    private _documentParser: DocumentParser;

    footnotes: WmlFootnote[]

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    parseXml(root: Element) {
        this.footnotes = this._documentParser.parseFootnotes(root);
    }
}