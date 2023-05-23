import { serializeXmlString } from "../parser/xml-parser";
import { OpenXmlPackage } from "./open-xml-package";
import { Relationship } from "./relationship";

export class Part {
    protected _xmlDocument: Document;

    rels: Relationship[];

    constructor(protected _package: OpenXmlPackage, public path: string) {
    }

    async load(): Promise<any> {
		this.rels = await this._package.loadRelationships(this.path);

		const xmlText = await this._package.load(this.path);
		const xmlDoc = this._package.parseXmlDocument(xmlText);

		if (this._package.options.keepOrigin) {
			this._xmlDocument = xmlDoc;
		}

		this.parseXml(xmlDoc.firstElementChild);
    }

    save() {
        this._package.update(this.path, serializeXmlString(this._xmlDocument));
    }

    protected parseXml(root: Element) {
    }
}