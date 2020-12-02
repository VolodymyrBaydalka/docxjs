import { Package } from "../common/package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IDomNumbering } from "../dom/dom";

export class NumberingPart extends Part {
    private _documentParser: DocumentParser;

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }

    numberings: IDomNumbering[];

    load(pkg: Package) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "string"))
            .then(xml => {
                this.numberings = this._documentParser.parseNumberingFile(xml);
            })
    }
}