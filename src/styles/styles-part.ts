import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IDomStyle } from "../document/style";

export class StylesPart extends Part {
    styles: IDomStyle[];

    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    parseXml(root: Element) {
        this.styles = this._documentParser.parseStylesFile(root);     
    }
}