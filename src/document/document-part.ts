import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { Relationship } from "../common/relationship";
import { DocumentParser } from "../document-parser";
import { DocumentElement } from "./document";

export class DocumentPart extends Part {
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    
    body: DocumentElement

    parseXml(root: Element) {
        this.body = this._documentParser.parseDocumentFile(root);
    }

    setParserExtraData(rels: Relationship[], chartPartsMap: Record<string, Part>) {
        this._documentParser.documentRels = rels;
        this._documentParser.chartPartsMap = chartPartsMap;
    }

    
}