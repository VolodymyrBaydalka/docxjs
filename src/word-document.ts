import * as JSZip from 'jszip';

import { DocumentParser } from './document-parser';
import { Relationship, RelationshipTypes } from './common/relationship';
import { Part } from './common/part';
import { FontTablePart } from './font-table/font-table';
import { Package } from './common/package';
import { DocumentPart } from './dom/document-part';
import { splitPath } from './utils';
import { NumberingPart } from './numbering/numbering-part';
import { StylesPart } from './styles/styles-part';

export class WordDocument {
    private _package: Package;
    private _parser: DocumentParser;
    
    rels: Relationship[];
    parts: Part[] = [];
    partsMap: Record<string, Part> = {};

    documentPart: DocumentPart;
    fontTablePart: FontTablePart;
    numberingPart: NumberingPart;
    stylesPart: StylesPart;

    static load(blob, parser: DocumentParser): Promise<WordDocument> {
        var d = new WordDocument();

        d._parser = parser;

        return JSZip.loadAsync(blob)
            .then(zip => {
                d._package = new Package(zip);

                return d._package.loadRelationships();
            }).then(rels => {
                d.rels = rels;

                let { target, type } = rels.find(x => x.type == RelationshipTypes.OfficeDocument) ?? {
                    target: "word/document.xml",
                    type: RelationshipTypes.OfficeDocument
                }; //fallback

                return d.loadRelationshipPart(target, type).then(() => d);
            });
    }

    private loadRelationshipPart(path: string, type: string): Promise<Part> {
        if (this.partsMap[path])
            return Promise.resolve(this.partsMap[path]);

        if (!this._package.exists(path))
            return Promise.resolve(null);

        let part: Part = null;

        switch(type) {
            case RelationshipTypes.OfficeDocument:
                this.documentPart = part = new DocumentPart(path, this._parser);
                break;

            case RelationshipTypes.FontTable:
                this.fontTablePart = part = new FontTablePart(path);
                break;

            case RelationshipTypes.Numbering:
                this.numberingPart = part = new NumberingPart(path, this._parser);
                break;

            case RelationshipTypes.Styles:
                this.stylesPart = part = new StylesPart(path, this._parser);
                break;
        }

        if (part == null)
            return Promise.resolve(null);

        this.partsMap[path] = part;
        this.parts.push(part);

        return part.load(this._package).then(() => {
            if (part.rels == null || part.rels.length == 0)
                return part;

            let [folder] = splitPath(part.path);
            let rels = part.rels.map(rel => {
                return this.loadRelationshipPart(`${folder}${rel.target}`, rel.type)
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

    private loadResource(part: Part, id: string, outputType: JSZip.OutputType) {
        let rel = part.rels.find(x => x.id == id);

        if (rel == null)
            return Promise.resolve(null);

        let [fodler] = splitPath(part.path);

        return this._package.load(fodler + rel.target, outputType);
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