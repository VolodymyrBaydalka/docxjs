import * as JSZip from 'jszip';

import { DocumentParser } from './document-parser';
import { IDomRelationship, IDomStyle, IDomNumbering } from './dom/dom';
import { Font } from './dom/common';
import { DocumentElement } from './dom/document';

enum PartType {
    Document = "word/document.xml",
    Style = "word/styles.xml",
    Numbering = "word/numbering.xml",
    FontTable = "word/fontTable.xml",
    DocumentRelations = "word/_rels/document.xml.rels",
    NumberingRelations = "word/_rels/numbering.xml.rels",
    FontRelations = "word/_rels/fontTable.xml.rels",
}

export class Document {
    private zip: JSZip = new JSZip();

    docRelations: IDomRelationship[] = null;
    fontRelations: IDomRelationship[] = null;
    numRelations: IDomRelationship[] = null;

    styles: IDomStyle[] = null;
    fonts: Font[] = null;
    fontTable: any;
    numbering: IDomNumbering[] = null;
    document: DocumentElement = null;

    static load(blob, parser: DocumentParser): PromiseLike<Document> {
        var d = new Document();

        return d.zip.loadAsync(blob).then(z => {
            var files = [
                d.loadPart(PartType.DocumentRelations, parser),
                d.loadPart(PartType.FontRelations, parser),
                d.loadPart(PartType.NumberingRelations, parser),
                d.loadPart(PartType.Style, parser),
                d.loadPart(PartType.FontTable, parser),
                d.loadPart(PartType.Numbering, parser),
                d.loadPart(PartType.Document, parser)
            ];

            return Promise.all(files.filter(x => x != null)).then(x => d);
        });
    }

    loadDocumentImage(id: string): PromiseLike<string> {
        return this.loadResource(this.docRelations, id, "blob")
            .then(x => x ? URL.createObjectURL(x) : null);
    }

    loadNumberingImage(id: string): PromiseLike<string> {
        return this.loadResource(this.numRelations, id, "blob")
            .then(x => x ? URL.createObjectURL(x) : null);
    }

    loadFont(id: string, key: string): PromiseLike<string> {
        return this.loadResource(this.fontRelations, id, "uint8array")
            .then(x => x ? URL.createObjectURL(new Blob([deobfuscate(x, key)])) : x);
    }

    private loadResource(relations: IDomRelationship[], id: string, outputType: JSZip.OutputType = "base64") {
        let rel = relations.find(x => x.id == id);
        return rel ? this.zip.files["word/" + rel.target].async(outputType) : Promise.resolve(null);
    }

    private loadPart(part: PartType, parser: DocumentParser) {
        var f = this.zip.files[part];

        return f ? f.async("text").then(xml => {
            switch (part) {
                case PartType.FontRelations:
                    this.fontRelations = parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.DocumentRelations:
                    this.docRelations = parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.NumberingRelations:
                    this.numRelations = parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.Style:
                    this.styles = parser.parseStylesFile(xml);
                    break;

                case PartType.Numbering:
                    this.numbering = parser.parseNumberingFile(xml);
                    break;

                case PartType.Document:
                    this.document = parser.parseDocumentFile(xml);
                    break;

                case PartType.FontTable:
                    this.fontTable = parser.parseFontTable(xml);
                    break;
            }

            return this;
        }) : null;
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