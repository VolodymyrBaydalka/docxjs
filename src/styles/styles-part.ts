import { Package } from "../common/package";
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

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }

    load(pkg: Package) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "xml"))
            .then(xml => {
                Object.assign(this, parseStylesPart(xml, pkg.xmlParser));
                this.domStyles = this._documentParser.parseStylesFile(xml);
            })
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