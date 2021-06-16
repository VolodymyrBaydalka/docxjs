import * as JSZip from "jszip";
import { parseXmlString, XmlParser } from "../parser/xml-parser";
import { splitPath } from "../utils";
import { parseRelationships, Relationship } from "./relationship";

export interface OpenXmlPackageOptions {
    trimXmlDeclaration: boolean,
    keepOrigin: boolean,
}

export class OpenXmlPackage {
    xmlParser: XmlParser = new XmlParser();

    constructor(private _zip: JSZip, public options: OpenXmlPackageOptions) {
    }

    get(path: string): any {
        return this._zip.files[normalizePath(path)];
    }

    update(path: string, content: any) {
        this._zip.file(path, content);
    }

    static load(input: Blob | any, options: OpenXmlPackageOptions): Promise<OpenXmlPackage> {
        return JSZip.loadAsync(input).then(zip => new OpenXmlPackage(zip, options));
    }

    save(type: any = "blob"): Promise<any>  {
        return this._zip.generateAsync({ type });
    }

    load(path: string, type: JSZip.OutputType = "string"): Promise<any> {
        return this.get(path)?.async(type) ?? Promise.resolve(null);
    }

    loadRelationships(path: string = null): Promise<Relationship[]> {
        let relsPath = `_rels/.rels`;

        if (path != null) {
            let [f, fn] = splitPath(path);
            relsPath = `${f}_rels/${fn}.rels`;
        }

        return this.load(relsPath)
            .then(txt => txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null);
    }

    /** @internal */
    parseXmlDocument(txt: string): Document {
        return parseXmlString(txt, this.options.trimXmlDeclaration);
    }
}

function normalizePath(path: string) {
    return path.startsWith('/') ? path.substr(1) : path;
}