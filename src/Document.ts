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
        Relations = "word/_rels/document.xml.rels"
    }

    export class Document {
        private zip: JSZip = new JSZip();

        relations: IDomRelationship[] = null;
        styles: IDomStyle[] = null;
        numbering: IDomNumbering[] = null;
        document: IDomDocument = null;

        static load(blob, parser: DocumentParser): PromiseLike<Document> {
            var d = new Document();
            
            return d.zip.loadAsync(blob).then(z => {
                var files = [d.loadPart(PartType.Relations, parser), d.loadPart(PartType.Style, parser), 
                    d.loadPart(PartType.Numbering, parser), d.loadPart(PartType.Document, parser)];

                return Promise.all(files.filter(x => x != null)).then(x => d);
            });
        }

        loadImage(id: string): PromiseLike<string> {
            var rel = this.relations.filter(x => x.id == id);

            if(rel.length == 0)
                return Promise.resolve(null);
            
            var file = this.zip.files["word/" + rel[0].target];

            return file.async("base64").then(x => "data:image/png;base64, " + x);
        }

        private loadPart(part: PartType, parser: DocumentParser) {
            var f = this.zip.files[part];

            return f ? f.async("string").then(xml => {
                switch(part) {
                    case PartType.Relations: 
                        this.relations = parser.parseDocumentRelationsFile(xml); 
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