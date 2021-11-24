import { Package } from "../common/package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { ParagraphElement } from "../dom/paragraph";

export class FooterPart extends Part {
    private _documentParser: DocumentParser;
    public paragraphs: ParagraphElement[] = [];
    public rootNode: Element;

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }

    load(pkg: Package) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "string"))
            .then(xml => {
                const result = this._documentParser.parseFooterFile(xml);
                this.paragraphs = result.content;
                this.rootNode = result.root;
            })
    }
}