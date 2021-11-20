import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DmlTheme, parseTheme } from "./theme";

export class ThemePart extends Part {
    theme: DmlTheme;

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

    parseXml(root: Element) {
        this.theme = parseTheme(root, this._package.xmlParser);
    }
}