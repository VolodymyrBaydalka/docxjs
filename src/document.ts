import * as JSZip from 'jszip';

import { DocumentParser } from './document-parser';
import { IDomRelationship, IDomStyle, IDomNumbering } from './dom/dom';
import { Font } from './dom/common';
import { DocumentElement } from './dom/document';
import { deobfuscate } from './deobfuscate';

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
        return this.loadResource(this.docRelations, id).then(x => x ? ("data:image/png;base64," + x) : null);
    }

    loadNumberingImage(id: string): PromiseLike<string> {
        return this.loadResource(this.numRelations, id).then(x => x ? ("data:image/png;base64," + x) : null);
    }

    loadFont(id: string, key: string): PromiseLike<string> {
        //const mimeType = "application/vnd.ms-package.obfuscated-opentype";
        const mimeType = "application/x-font-ttf";
        //const mimeType = "application/vnd.ms-fontobject";
        return this.loadResource(this.fontRelations, id, "array")
            .then(x => x ? deobfuscate(x, key) : x);
            //.then(x => x ? (`data:${mimeType};base64,${x}`) : null);
    }

    private loadResource(relations: IDomRelationship[], id: string, output: "base64" | "array" = "base64") {
        let rel = relations.filter(x => x.id == id);

        return rel.length == 0 ? Promise.resolve(null) : this.zip.files["word/" + rel[0].target].async(output);
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
