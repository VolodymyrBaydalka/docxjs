
namespace docx {
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

        parseNumberingAsync(zip){
            var file = zip.files["word/numbering.xml"];
            return file ? file.async("string")
                .then((xml) => this.parseNumberingFile(xml)) : null;
        }

        parseDocumentFile(xmlString) {
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
                target: null,
                styles: []
            };

            switch (xml.stringAttr(node, "type")) {
                case "paragraph": result.target = "p"; break;
                case "table": result.target = "table"; break;
            }

            xml.foreach(node, n => {
                switch (n.localName) {
                    case "pPr":
                        result.styles.push({
                            target: "p",
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;

                    case "rPr":
                        result.styles.push({
                            target: "span",
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;

                    case "tblPr":
                        result.styles.push({
                            target: "td", //TODO: maybe move to processor
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                }
            });

            return result;
        }

        parseNumberingFile(xmlString: string): IDomStyle[] {
            var result = [];
            var xnums = xml.parse(xmlString, this.skipDeclaration);

            xml.foreach(xnums, n => {
            });

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
                        this.parseParagraphProperties(c, result);
                        break;
                }
            });

            return result;
        }

        parseParagraphProperties(node: Node, paragraph: IDomParagraph) {
            paragraph.style = this.parseDefaultProperties(node, {}, null, c => {
                switch (c.localName) {
                    case "pStyle":
                        paragraph.className = xml.stringAttr(c, "val");
                        break;
                    
                    case "numPr":

                        break;

                    default:
                        return false;
                }

                return true;
            });
        }

        parseNumbering(node: Node, paragraph: IDomParagraph){
             xml.foreach(node, c => {
                switch (c.localName) {
                    case "numId":
                        paragraph.numberingId = xml.stringAttr(c, "val");
                        break;

                    case "ilvl":
                        paragraph.numberingLevel = xml.stringAttr(c, "val");
                        break;
                }
            });
        }

        parseRun(node: Node): IDomElement {
            var result: IDomRun = { domType: DomType.Run };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "t":
                        result.text = c.textContent;//.replace(" ", "\u00A0"); // TODO
                        break;

                    case "br":
                        result.break = xml.stringAttr(c, "type") || "textWrapping";
                        break;

                    case "rPr":
                        result.style = this.parseDefaultProperties(c, {}, null);
                        break;
                }
            });

            return result;
        }

        parseTable(node: Node): IDomTable {
            var result: IDomTable = { domType: DomType.Table, children: [] };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "tr":
                        result.children.push(this.parseTableRow(c));
                        break;

                    case "tblPr":
                        this.parseTableProperties(c, result);
                        break;
                }
            });

            return result;
        }

        parseTableProperties(node: Node, table: IDomTable) {
            table.style = {};
            table.cellStyle = {};

            this.parseDefaultProperties(node, table.style, table.cellStyle, c => {
                switch (c.localName) {
                    case "tblStyle":
                        table.className = xml.stringAttr(c, "val");
                        break;

                    default:
                        return false;
                }

                return true;
            });
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
            var result: IDomTableCell = { domType: DomType.Cell, children: [] };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "tbl":
                        result.children.push(this.parseTable(c));
                        break;

                    case "p":
                        result.children.push(this.parseParagraph(c));
                        break;

                    case "tcPr":
                        this.parseTableCellProperties(c, result);
                        break;
                }
            });

            return result;
        }

        parseTableCellProperties(node: Node, cell: IDomTableCell) {
            cell.style = this.parseDefaultProperties(node, {}, null, c => {
                switch (c.localName) {
                    case "gridSpan":
                        cell.span = xml.intAttr(c, "val", null);
                        break;

                    case "vAlign":
                        cell.vAlign = xml.stringAttr(c, "val");
                        break;

                    default:
                        return false;
                }

                return true;
            });
        }

        parseDefaultProperties(node: Node, style: IDomStyleValues = null, childStyle: IDomStyleValues = null, handler: (prop: Node) => void = null): IDomStyleValues {
            xml.foreach(node, c => {
                switch (c.localName) {
                    case "jc":
                        style["text-align"] = values.valueOfJc(c);
                        break;

                    case "color":
                        style["color"] = xml.stringAttr(c, "val");
                        break;
                    
                    case "sz":
                        style["font-size"] = xml.sizeAttr(c, "val", SizeType.FontSize);
                        break;

                    case "shd":
                        style["background-color"] = xml.stringAttr(c, "fill");
                        break;

                    case "highlight":
                        style["background-color"] = xml.stringAttr(c, "val");
                        break;

                    case "strike":
                        style["text-decoration"] = values.valueOfStrike(c);
                        break;

                    case "b":
                        style["font-weight"] = values.valueOfBold(c);
                        break;

                    case "i":
                        style["font-style"] = "italic";
                        break;

                    case "u":
                        style["text-decoration"] = "underline";
                        break;

                    case "ind":
                        style["text-indent"] = xml.sizeAttr(c, "left");
                        break;

                    case "tblBorders":
                        this.parseBorderProperties(c, childStyle || style);
                        break;

                    case "tcBorders":
                        this.parseBorderProperties(c, style);
                        break;

                    case "tblCellMar":
                        this.parseMarginProperties(c, childStyle || style);
                        break;

                    case "tblLayout":
                        style["table-layout"] = values.valueOfTblLayout(c);
                        break;

                    case "spacing":
                        this.parseSpacing(c, style);
                        break;

                    case "lang":
                        break;

                    default:
                        if (handler == null || !handler(c))
                            console.log(c.localName);
                        break;
                }
            });

            return style;
        }

        parseSpacing(node: Node, style: IDomStyleValues) {
            var before = xml.sizeAttr(node, "before");
            var after = xml.sizeAttr(node, "after");
            var line = xml.sizeAttr(node, "line");

            if(before) style["magrin-top"] = before;
            if(after) style["magrin-bottom"] = after;
            if(line) style["line-height"] = line;
        }

        parseMarginProperties(node: Node, output: IDomStyleValues) {
            xml.foreach(node, c => {
                switch (c.localName) {
                    case "left":
                        output["padding-left"] = values.valueOfMargin(c);
                        break;

                    case "right":
                        output["padding-right"] = values.valueOfMargin(c);
                        break;

                    case "top":
                        output["padding-top"] = values.valueOfMargin(c);
                        break;

                    case "bottom":
                        output["padding-bottom"] = values.valueOfMargin(c);
                        break;
                }
            });
        }

        parseBorderProperties(node: Node, output: IDomStyleValues) {
            xml.foreach(node, c => {
                switch (c.localName) {
                    case "left":
                        output["border-left"] = values.valueOfBorder(c);
                        break;

                    case "right":
                        output["border-right"] = values.valueOfBorder(c);
                        break;

                    case "top":
                        output["border-top"] = values.valueOfBorder(c);
                        break;

                    case "bottom":
                        output["border-bottom"] = values.valueOfBorder(c);
                        break;
                }
            });
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

        static boolAttr(node: Node, attrName: string, defValue: boolean = false) {
            var v = xml.stringAttr(node, attrName);

            switch (v)
            {
                case "1": return true;
                case "0": return false;
            }

            return defValue;
        }

        static intAttr(node: Node, attrName: string, defValue: number = 0) {
            var val = xml.stringAttr(node, attrName);
            return val ? parseInt(xml.stringAttr(node, attrName)) : 0;
        }

        static sizeAttr(node: Node, attrName: string, type: SizeType = SizeType.Dxa) {
            var val = xml.stringAttr(node, attrName);

            if (val == null || val.indexOf("pt") > -1)
                return val;

            var intVal = parseInt(val);

            switch (type) {
                case SizeType.Dxa: return (0.05 * intVal) + "pt";
                case SizeType.FontSize: return (0.5 * intVal) + "pt";
                case SizeType.Percent: return (0.01 * intVal) + "%";
            }

            return val;
        }
    }

    class values {
        static valueOfBold(c: Node) {
            return xml.boolAttr(c, "val", true) ? "bold" : "normal"
        }

        static valueOfStrike(c: Node) {
            return xml.boolAttr(c, "val", true) ? "line-through" : "none"
        }

        static valueOfMargin(c: Node) {
            return xml.sizeAttr(c, "w");
        }

        static valueOfBorder(c: Node) {
            var type = xml.stringAttr(c, "val");

            if (type == "nil")
                return "none";

            var color = xml.stringAttr(c, "color");
            var size = xml.sizeAttr(c, "sz");

            return `${size} solid ${color == "auto" ? "black" : color}`;
        }

        static valueOfTblLayout(c: Node) {
            var type = xml.stringAttr(c, "val");
            return type == "fixed" ? "fixed" : "auto";
        }

        static valueOfJc(c: Node) {
            var type = xml.stringAttr(c, "val");

            switch(type){
                case "start": 
                case "left": return "left";
                case "center": return "center";
                case "end": 
                case "right": return "right";
                case "both": return "justify";
            }

            return type;
        }
    }
}
