import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IDomStyle } from "../dom/style";
import { XmlParser } from "../parser/xml-parser";
import { DocumentDefaults, parseDocumentDefaults } from "./document-defaults";
import { parseStyle, Style, StyleType } from "./style";

export class StylesPart extends Part implements StylesPartProperties {
    defaults: DocumentDefaults;
    styles: Style[];
    domStyles: IDomStyle[];

    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    parseXml(root: Element) {
        Object.assign(this, parseStylesPart(root, this._package.xmlParser));
        this.domStyles = this._documentParser.parseStylesFile(root);     
    }
}

export interface StylesPartProperties {
    defaults: DocumentDefaults;
    styles: Style[];
}

export function parseStylesPart(elem: Element, xml: XmlParser): StylesPartProperties {
    let result = {
        styles: []
    } as StylesPartProperties;

    for (let e of xml.elements(elem)) {
        switch (e.localName) {
            case "docDefaults":
                result.defaults = parseDocumentDefaults(e, xml);
                break;

            case "style":
                result.styles.push(parseStyle(e, xml));
                break;
        }
    }

    return result;
}