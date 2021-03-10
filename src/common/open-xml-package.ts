import * as JSZip from "jszip";
import { parseXmlString, XmlParser } from "../parser/xml-parser";
import { splitPath } from "../utils";
import { parseRelationships, Relationship } from "./relationship";

export class OpenXmlPackage {
    xmlParser: XmlParser = new XmlParser();

    constructor(private _zip: JSZip) {
    }

    exists(path: string): boolean {
        return this._zip.files[path] != null;
    }

    update(path: string, content: any) {
        this._zip.file(path, content);
    }

    static load(input: Blob | any): Promise<OpenXmlPackage> {
        return JSZip.loadAsync(input).then(zip => new OpenXmlPackage(zip));
    }

    save(type: any = "blob"): Promise<any>  {
        return this._zip.generateAsync({ type });
    }

    load(path: string, type: JSZip.OutputType): Promise<any> {
        return this._zip.files[path]?.async(type) ?? Promise.resolve(null);
    }

    loadRelationships(path: string = null): Promise<Relationship[]> {
        let relsPath = `_rels/.rels`;

        if (path != null) {
            let [f, fn] = splitPath(path);
            relsPath = `${f}_rels/${fn}.rels`;
        }

        return this.load(relsPath, "string").then(text => {
            if (!text)
                return null;

            return parseRelationships(parseXmlString(text).firstElementChild, this.xmlParser);
        })
    }
}