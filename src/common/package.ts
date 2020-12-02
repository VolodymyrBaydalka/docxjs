import JSZip = require("jszip");
import { XmlParser } from "../parser/xml-parser";
import { splitPath } from "../utils";
import { parseRelationships, Relationship } from "./relationship";

export class Package {
    xmlParser: XmlParser = new XmlParser();

    constructor(private _zip: JSZip) {
    }

    exists(path: string): boolean {
        return this._zip.files[path] != null;
    }

    load(path: string, type: "xml" | JSZip.OutputType): Promise<any> {
        let file = this._zip.files[path];

        if (file == null)
            return Promise.resolve(null);

        if (type == "xml")
            return file.async("string").then(t => this.xmlParser.parse(t));

        return file.async(type);
    }

    loadRelationships(path: string = null): Promise<Relationship[]> {
        let relsPath = `_rels/.rels`;

        if (path != null) {
            let [f, fn] = splitPath(path);
            relsPath = `${f}_rels/${fn}.rels`;
        }

        return this.load(relsPath, "xml").then(xml => {
            return xml == null ? null : parseRelationships(xml, this.xmlParser);
        })
    }
}