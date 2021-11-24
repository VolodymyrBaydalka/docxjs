import { OutputType } from "jszip";

import { DocumentParser } from './document-parser';
import { Relationship, RelationshipTypes } from './common/relationship';
import { Part } from './common/part';
import { FontTablePart } from './font-table/font-table';
import { OpenXmlPackage } from './common/open-xml-package';
import { DocumentPart } from './document/document-part';
import { resolvePath, splitPath } from './utils';
import { NumberingPart } from './numbering/numbering-part';
import { StylesPart } from './styles/styles-part';
import { FooterPart } from "./footer/footer-part";
import { HeaderPart } from "./header/header-part";
import { ExtendedPropsPart } from "./document-props/extended-props-part";
import { CorePropsPart } from "./document-props/core-props-part";
import { ThemePart } from "./theme/theme-part";
import { FootnotesPart } from "./footnotes/footnotes-part";

const topLevelRels = [
    { type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
    { type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
    { type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
];

export class WordDocument {
    private _package: OpenXmlPackage;
    private _parser: DocumentParser;
    
    rels: Relationship[];
    parts: Part[] = [];
    partsMap: Record<string, Part> = {};

    documentPart: DocumentPart;
    fontTablePart: FontTablePart;
    numberingPart: NumberingPart;
    stylesPart: StylesPart;
    footnotesPart: FootnotesPart;
    corePropsPart: CorePropsPart;
    extendedPropsPart: ExtendedPropsPart;

    static load(blob, parser: DocumentParser, options: any): Promise<WordDocument> {
        var d = new WordDocument();

        d._parser = parser;

        return OpenXmlPackage.load(blob, options)
            .then(pkg => {
                d._package = pkg;

                return d._package.loadRelationships();
            }).then(rels => {
                d.rels = rels;

                const tasks = topLevelRels.map(rel => {
                    const r = rels.find(x => x.type === rel.type) ?? rel; //fallback                    
                    return d.loadRelationshipPart(r.target, r.type);
                });

                return Promise.all(tasks);
            }).then(() => d);
    }

    save(type = "blob"): Promise<any> {
        return this._package.save(type);
    }

    private loadRelationshipPart(path: string, type: string): Promise<Part> {
        if (this.partsMap[path])
            return Promise.resolve(this.partsMap[path]);

        if (!this._package.get(path))
            return Promise.resolve(null);

        let part: Part = null;

        switch(type) {
            case RelationshipTypes.OfficeDocument:
                this.documentPart = part = new DocumentPart(this._package, path, this._parser);
                break;

            case RelationshipTypes.FontTable:
                this.fontTablePart = part = new FontTablePart(this._package, path);
                break;

            case RelationshipTypes.Numbering:
                this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
                break;

            case RelationshipTypes.Styles:
                this.stylesPart = part = new StylesPart(this._package, path, this._parser);
                break;

            case RelationshipTypes.Theme:
                part = new ThemePart(this._package, path);
                break;

            case RelationshipTypes.Footnotes:
                this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
                break;
        
            case RelationshipTypes.Footer:
                part = new FooterPart(this._package, path, this._parser);
                break;

            case RelationshipTypes.Header:
                part = new HeaderPart(this._package, path, this._parser);
                break;

            case RelationshipTypes.CoreProperties:
                this.corePropsPart = part = new CorePropsPart(this._package, path);
                break;

            case RelationshipTypes.ExtendedProperties:
                this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
                break;
        }

        if (part == null)
            return Promise.resolve(null);

        this.partsMap[path] = part;
        this.parts.push(part);

        return part.load().then(() => {
            if (part.rels == null || part.rels.length == 0)
                return part;

            const [folder] = splitPath(part.path); 
            const rels = part.rels.map(rel => {
                return this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)
            });

            return Promise.all(rels).then(() => part);
        });
    }

    loadDocumentImage(id: string): PromiseLike<string> {
        return this.loadResource(this.documentPart, id, "blob")
            .then(x => x ? URL.createObjectURL(x) : null);
    }

    loadNumberingImage(id: string): PromiseLike<string> {
        return this.loadResource(this.numberingPart, id, "blob")
            .then(x => x ? URL.createObjectURL(x) : null);
    }

    loadFont(id: string, key: string): PromiseLike<string> {
        return this.loadResource(this.fontTablePart, id, "uint8array")
            .then(x => x ? URL.createObjectURL(new Blob([deobfuscate(x, key)])) : x);
    }

    findPartByRelId(id: string, basePart: Part = null) {
        var rel = (basePart.rels ?? this.rels).find(r => r.id == id);
        const folder = basePart ? splitPath(basePart.path)[0] : ''; 
        return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
    }

    getPathById(part: Part, id: string): string {
        const rel = part.rels.find(x => x.id == id);
        const [folder] = splitPath(part.path); 
        return rel ? resolvePath(rel.target, folder) : null;
    }

    private loadResource(part: Part, id: string, outputType: OutputType) {
        const path = this.getPathById(part, id);
        return path ? this._package.load(path, outputType) : Promise.resolve(null);
    }
}

export function deobfuscate(data: Uint8Array, guidKey: string): Uint8Array {
    const len = 16;
    const trimmed = guidKey.replace(/{|}|-/g, "");
    const numbers = new Array(len);
    
    for(let i = 0; i < len; i ++)
        numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);

    for (let i = 0; i < 32; i++)
        data[i] = data[i] ^ numbers[i % len]

    return data;
}