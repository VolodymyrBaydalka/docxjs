import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { FooterElement } from "./footer";

export class FooterPart extends Part {
    footerElement: FooterElement;
    
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    
    parseXml(root: Element) {
        this.footerElement = this._documentParser.parseFooter(root);
    }
}