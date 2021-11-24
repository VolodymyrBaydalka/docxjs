import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { WmlHeader } from "./header";

export class HeaderPart extends Part {
    headerElement: WmlHeader;

    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    parseXml(root: Element) {
        this.headerElement = this._documentParser.parseHeader(root);
    }
}