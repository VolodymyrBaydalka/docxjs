import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { FontDeclaration, parseFonts } from "./fonts";

export class FontTablePart extends Part {
    fonts: FontDeclaration[];

    load(pkg: OpenXmlPackage): Promise<void> {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "xml"))
            .then(el => {
                this.fonts = parseFonts(el, pkg.xmlParser);
            });
    }
}