
module docx {
    export class DocumentParser {
        public skipDeclaration: boolean = true;

        parseDocumentAsync(zip) {
            return zip.files["word/document.xml"]
                .async("string")
                .then((xml) => this.parseDocumentFile(xml));
        }

        parseStylesAsync(zip) {
            return zip.files["word/styles.xml"]
                .async("string")
                .then((xml) => this.parseStylesFile(xml));
        }

        private parseDocumentFile(xmlString) {
            var result: IDomDocument = {
                domType: DomType.Document,
                children: [],
                style: {}
            };

            var xbody = xml.byTagName(xml.parse(xmlString, this.skipDeclaration), "body");

            for (var i = 0; i < xbody.childNodes.length; i++) {
                var node = xbody.childNodes[i];

                switch (node.localName) {
                    case "p":
                        result.children.push(this.parseParagraph(node));
                        break;

                    case "tbl":
                        result.children.push(this.parseTable(node));
                        break;

                    case "sectPr":
                        this.parseSectionProperties(node, result);
                        break;
                }
            }

            return result;
        }

        parseStylesFile(xmlString: string): IDomStyle[] {
            var result = [];

            var xstyles = xml.parse(xmlString, this.skipDeclaration);

            xml.foreach(xstyles, n => {
                switch (n.localName) {
                    case "style":
                        result.push(this.parseStyle(n));
                        break;
                }
            });

            return result;
        }

        parseStyle(node: Node): IDomStyle {
            var result = {
                id: xml.stringAttr(node, "styleId"),
                target: null
            };

            switch (xml.stringAttr(node, "type")) {
                case "paragraph": result.target = "p"; break;
                case "table": result.target = "table"; break;
            }

            return result;
        }

        parseSectionProperties(node: Node, elem: IDomElement) {
            xml.foreach(node, n => {
                switch (n.localName) {
                    case "pgMar":
                        elem.style["padding-left"] = xml.sizeAttr(n, "left");
                        elem.style["padding-right"] = xml.sizeAttr(n, "right");
                        elem.style["padding-top"] = xml.sizeAttr(n, "top");
                        elem.style["padding-bottom"] = xml.sizeAttr(n, "bottom");
                        break;
                }
            });
        }

        parseParagraph(node: Node): IDomElement {
            var result: IDomElement = { domType: DomType.Paragraph, children: [] };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "r":
                        result.children.push(this.parseRun(c));
                        break;

                    case "pPr":
                        result.style = this.parseDefaultProperties(c, {}, null);
                        break;
                }
            });

            return result;
        }

        parseRun(node: Node): IDomElement {
            var result: IDomRun = { domType: DomType.Run };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "t":
                        result.text = c.textContent;
                        break;

                    case "br":
                        result.isBreak = true;
                        break;

                    case "rPr":
                        result.style = this.parseDefaultProperties(c, {}, null);
                        break;
                }
            });

            return result;
        }

        parseTable(node: Node): IDomElement {
            var result: IDomElement = { domType: DomType.Table, children: [] };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "tr":
                        result.children.push(this.parseTableRow(c));
                        break;
                }
            });

            return result;
        }

        parseTableRow(node: Node): IDomElement {
            var result: IDomElement = { domType: DomType.Row, children: [] };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "tc":
                        result.children.push(this.parseTableCell(c));
                        break;
                }
            });

            return result;
        }
        parseTableCell(node: Node): IDomElement {
            var result: IDomElement = { domType: DomType.Cell, children: [] };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "tbl":
                        result.children.push(this.parseTable(c));
                        break;

                    case "p":
                        result.children.push(this.parseParagraph(c));
                        break;
                }
            });

            return result;
        }

        parseDefaultProperties(node: Node, output: IDomStyleValues, handler: (prop: Node) => void): IDomStyleValues {
            xml.foreach(node, c => {
                switch (c.localName) {
                    case "jc":
                        output["text-align"] = xml.stringAttr(c, "val");
                        break;

                    case "color":
                        output["color"] = xml.stringAttr(c, "val");
                        break;

                    case "b":
                        output["font-weight"] = "bold";
                        break;

                    case "i":
                        output["font-style"] = "italic";
                        break;

                    default:
                        if (handler != null)
                            handler(node);
                        break;
                }
            });

            return output;
        }
    }

    enum SizeType {
        FontSize,
        Dxa,
        Percent
    }

    class xml {
        static parse(xmlString, skipDeclaration = true) {
            if (skipDeclaration)
                xmlString = xmlString.replace(/<[?].*[?]>/, "");

            return new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
        }

        static foreach(node: Node, cb: (n: Node) => void) {
            for (var i = 0; i < node.childNodes.length; i++)
                cb(node.childNodes.item(i));
        }

        static byTagName(node: Node, tagName: string) {
            for (var i = 0; i < node.childNodes.length; i++)
                if (node.childNodes[i].localName == tagName)
                    return node.childNodes[i];
        }

        static stringAttr(node: Node, attrName: string) {
            for (var i = 0; i < node.attributes.length; i++) {
                var attr = node.attributes.item(i);

                if (attr.localName == attrName)
                    return attr.value;
            }

            return null;
        }

        static sizeAttr(node: Node, attrName: string, type: SizeType = SizeType.Dxa) {
            var val = xml.stringAttr(node, attrName);

            var intVal = parseInt(val);

            switch (type) {
                case SizeType.Dxa: return (0.05 * intVal) + "pt";
                case SizeType.FontSize: return (0.5 * intVal) + "pt";
                case SizeType.Percent: return (0.01 * intVal) + "%";
            }

            return val;
        }
    }
}
