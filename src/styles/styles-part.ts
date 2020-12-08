import { Package } from "../common/package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IDomStyle } from "../dom/style";

export class StylesPart extends Part {
    styles: IDomStyle[];
    
    private _documentParser: DocumentParser;

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }

    load(pkg: Package) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "string"))
            .then(xml => {
                this.styles = this._documentParser.parseStylesFile(xml);
            })
    }
}