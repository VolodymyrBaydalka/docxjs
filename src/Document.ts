namespace docx {
    export interface Options {
        inWrapper: boolean;
        ignoreWidth: boolean;
        ignoreHeight: boolean;
        debug: boolean;
        className: string;
    }

    export function renderAsync(data, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options = null): PromiseLike<any> {
        var parser = new docx.DocumentParser();
        var renderer = new docx.HtmlRenderer(window.document);

        if (options) {
            parser.ignoreWidth = options.ignoreWidth || parser.ignoreWidth;
            parser.ignoreHeight = options.ignoreHeight || parser.ignoreHeight;
            parser.debug = options.debug || parser.debug;

            renderer.className = options.className || "docx";
            renderer.inWrapper = options && options.inWrapper != null ? options.inWrapper : true;
        }

        return Document.load(data, parser)
            .then(doc => {
                renderer.render(doc, bodyContainer, styleContainer);
                return doc;
            });
    }

    enum PartType {
        Document = "word/document.xml",
        Style = "word/styles.xml",
        Numbering = "word/numbering.xml",
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
        fonts: IDomFont[] = null;
        numbering: IDomNumbering[] = null;
        document: IDomDocument = null;

        static load(blob, parser: DocumentParser): PromiseLike<Document> {
            var d = new Document();
            
            return d.zip.loadAsync(blob).then(z => {
                var files = [
                    d.loadPart(PartType.DocumentRelations, parser), 
                    d.loadPart(PartType.FontRelations, parser), 
                    d.loadPart(PartType.NumberingRelations, parser), 
                    d.loadPart(PartType.Style, parser), 
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

        loadFont(id: string): PromiseLike<string> {
            return this.loadResource(this.fontRelations, id)
                .then(x => x ? ("data:application/vnd.ms-package.obfuscated-opentype;charset=utf-8;base64," + x) : null);
        }

        private loadResource(relations: IDomRelationship[], id: string) {
            let rel = relations.filter(x => x.id == id);

            return rel.length == 0 ? Promise.resolve(null) : this.zip.files["word/" + rel[0].target].async("base64");
        }

        private loadPart(part: PartType, parser: DocumentParser) {
            var f = this.zip.files[part];

            return f ? f.async("string").then(xml => {
                switch(part) {
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
                }

                return this;
            }) : null;
        }
    }
}