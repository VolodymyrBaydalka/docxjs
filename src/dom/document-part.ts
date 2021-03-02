import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { DocumentElement } from "./document";

export class DocumentPart extends Part {
    private _documentParser: DocumentParser;

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }
    
    documentElement: DocumentElement

    load(pkg: OpenXmlPackage) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "xml"))
            .then(xml => {
                this.documentElement = this._documentParser.parseDocumentFile(xml);
            });
    }
}