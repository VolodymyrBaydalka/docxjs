import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { OpenXmlElement } from "../document/dom";
import { WmlHeader, WmlFooter } from "./elements";

export abstract class BaseHeaderFooterPart<T extends OpenXmlElement = OpenXmlElement> extends Part {
    rootElement: T;

    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    parseXml(root: Element) {
        this.rootElement = this.createRootElement();
        this.rootElement.children = this._documentParser.parseBodyElements(root);
    }

    protected abstract createRootElement(): T;
}

export class HeaderPart extends BaseHeaderFooterPart<WmlHeader> {
    protected createRootElement(): WmlHeader {
        return new WmlHeader();
    }
}

export class FooterPart extends BaseHeaderFooterPart<WmlFooter> {
    protected createRootElement(): WmlFooter {
        return new WmlFooter();
    }
}