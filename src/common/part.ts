import { parseXmlString, serializeXmlString } from "../parser/xml-parser";
import { OpenXmlPackage } from "./open-xml-package";
import { Relationship } from "./relationship";

export class Part {
    protected _xmlDocument: Document;

    rels: Relationship[];

    constructor(protected _package: OpenXmlPackage, public path: string) {
    }

    load(): Promise<any> {
        return Promise.all([
            this._package.loadRelationships(this.path).then(rels => {
                this.rels = rels;
            }),
            this._package.load(this.path, 'string').then(text => {
                this._xmlDocument = parseXmlString(text);
                this.parseXml(this._xmlDocument.firstElementChild);
            })
        ]);
    }

    save() {
        this._package.update(this.path, serializeXmlString(this._xmlDocument));
    }

    protected parseXml(root: Element) {
    }
}