import { Package } from "../common/package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IDomNumbering } from "../dom/dom";
import { AbstractNumbering, Numbering, NumberingBulletPicture, NumberingPartProperties, parseNumberingPart } from "./numbering";

export class NumberingPart extends Part implements NumberingPartProperties {
    private _documentParser: DocumentParser;

    constructor(path: string, parser: DocumentParser) {
        super(path);
        this._documentParser = parser;
    }

    numberings: Numbering[];
    abstractNumberings: AbstractNumbering[];
    bulletPictures: NumberingBulletPicture[];
    
    domNumberings: IDomNumbering[];

    load(pkg: Package) {
        return super.load(pkg)
            .then(() => pkg.load(this.path, "xml"))
            .then(xml => {
                Object.assign(this, parseNumberingPart(xml, pkg.xmlParser));
                this.domNumberings = this._documentParser.parseNumberingFile(xml);
            })
    }
}