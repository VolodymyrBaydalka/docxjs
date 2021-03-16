import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { WmlDocument } from "./document";

export class DocumentPart extends Part {
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    
    documentElement: WmlDocument

    parseXml(root: Element) {
        this.documentElement = this._documentParser.parseDocumentFile(root);
    }
}