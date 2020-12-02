import { Package } from "../common/package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { DocumentElement } from "./document";

export class DocumentPart extends Part {
    private _documentParser: DocumentParser;

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }
    
    body: DocumentElement

    load(pkg: Package) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "string"))
            .then(xml => {
                this.body = this._documentParser.parseDocumentFile(xml);
            });
    }
}