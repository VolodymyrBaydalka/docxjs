import { Package } from "../common/package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import {ImportantFonts} from "../font-table/fonts";

export class ThemesPart extends Part {
    importantFonts: ImportantFonts;

    private _documentParser: DocumentParser;

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }

    load(pkg: Package) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "string"))
            .then(xml => {
                this.importantFonts = this._documentParser.parseThemesFile(xml);
            })
    }
}