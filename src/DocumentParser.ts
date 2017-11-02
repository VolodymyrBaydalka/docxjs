
namespace docx {
    export class DocumentParser {
        // removes XML declaration 
        skipDeclaration: boolean = true;
        
         // ignores page and table sizes
        ignoreWidth: boolean = false;
        ignoreHeight: boolean = true; 
        debug: boolean = false;

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

        parseDocumentRelationsAsync(zip){
            var file = zip.files["word/_rels/document.xml.rels"]
            return file ? file.async("string")
                .then((xml) => this.parseDocumentRelationsFile(xml)) : null;
        }

        parseDocumentRelationsFile(xmlString) {
            var xrels = xml.parse(xmlString, this.skipDeclaration);

            return xml.nodes(xrels).map(c => <IDomRelationship>{
                id: xml.stringAttr(c, "Id"),
                type: values.valueOfRelType(c),
                target: xml.stringAttr(c, "Target"),
            });
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

                    case "docDefaults":
                        result.push(this.parseDefaultStyles(n));
                        break;
                }
            });

            return result;
        }

        parseDefaultStyles(node: Node): IDomStyle {
            var result = {
                id: null,
                name: null,
                target: null,
                basedOn: null,
                styles: []
            };

            xml.foreach(node, c => {
                switch(c.localName) {
                    case "rPrDefault": 
                        var rPr = xml.byTagName(c, "rPr");
                        
                        if(rPr)
                            result.styles.push({
                                target: "span",
                                values: this.parseDefaultProperties(rPr, {})
                            });
                        break;

                    case "pPrDefault": 
                        var pPr = xml.byTagName(c, "pPr");

                        if(pPr)
                            result.styles.push({
                                target: "p",
                                values: this.parseDefaultProperties(pPr, {})
                            });
                        break;
                }
            });

            return result;
        }

        parseStyle(node: Node): IDomStyle {
            var result = <IDomStyle>{
                id: xml.stringAttr(node, "styleId"),
                isDefault: xml.boolAttr(node, "default"),
                name: null,
                target: null,
                basedOn: null,
                styles: []
            };

            switch (xml.stringAttr(node, "type")) {
                case "paragraph": result.target = "p"; break;
                case "table": result.target = "table"; break;
            }

            xml.foreach(node, n => {
                switch (n.localName) {
                    case "basedOn":
                        result.basedOn = xml.stringAttr(n, "val");
                        break;

                    case "name":
                        result.name = xml.stringAttr(n, "val");
                        break;

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
                    case "tcPr":
                        result.styles.push({
                            target: "td", //TODO: maybe move to processor
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;

                    case "tblStylePr":
                        for(let s of this.parseTableStyle(n))
                            result.styles.push(s);
                        break;

                    case "rsid":
                    case "qFormat":
                    case "semiHidden":
                    case "uiPriority":
                        //TODO: ignore
                        break;
    
                    default:
                    this.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
                }
            });

            return result;
        }

        parseTableStyle(node: Node): IDomSubStyle[] {
            var result = [];

            var type = xml.stringAttr(node, "type");
            var selector = "";

            switch(type){
                case "firstRow": selector = "tr.first-row"; break;
                case "lastRow": selector = "tr.last-row"; break;
                case "firstCol": selector = "td.first-col"; break;
                case "lastCol": selector = "td.first-col"; break;
                case "band1Vert": selector = "td.odd-col"; break;
                case "band2Vert": selector = "td.even-col"; break;
                case "band1Horz": selector = "tr.odd-row"; break;
                case "band2Horz": selector = "tr.even-row"; break;
                default: return [];
            }

            xml.foreach(node, n => {
                switch (n.localName) {
                    case "pPr":
                        result.push({
                            target: selector + " p",
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;

                    case "rPr":
                        result.push({
                            target: selector + " span",
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;

                    case "tblPr":
                    case "tcPr":
                        result.push({
                            target: selector, //TODO: maybe move to processor
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                }
            });

            return result;
        }

        parseNumberingFile(xmlString: string): IDomNumbering[] {
            var result = [];
            var xnums = xml.parse(xmlString, this.skipDeclaration);
            
            var mapping = {};

            xml.foreach(xnums, n => {
                switch(n.localName){
                    case "abstractNum":
                        this.parseAbstractNumbering(n)
                            .forEach(x => result.push(x));
                        break;

                    case "num":
                        var numId = xml.stringAttr(n, "numId");
                        var abstractNumId = xml.nodeStringAttr(n, "abstractNumId", "val");
                        mapping[abstractNumId] = numId;
                        break;
                }
            });

            result.forEach(x => x.id = mapping[x.id]);

            return result;
        }

        parseAbstractNumbering(node: Node): IDomNumbering[] {
            var result = [];
            var id = xml.stringAttr(node, "abstractNumId"); 

            xml.foreach(node, n => {
                switch (n.localName) {
                    case "lvl":
                        result.push({
                            id: id,
                            level:  xml.stringAttr(n, "ilvl"),
                            style: this.parseNumberingLevel(n)
                        });
                        break;
                }
            });

            return result;
        }

	    parseNumberingLevel(node: Node): IDomStyleValues {
            var result = <IDomStyleValues>{}; 

            xml.foreach(node, n => {
                switch (n.localName) {
                    case "pPr":
                        this.parseDefaultProperties(n, result);
                        break;
                    
                    case "lvlText":
                        break;

                    case "numFmt":
                        this.parseNumberingFormating(n, result);
                        break;
                }
            });

            return result;
        }

        parseNumberingFormating(node: Node, style: IDomStyleValues) {
            switch(xml.stringAttr(node, "val")) 
            {
                case "bullet":
                    style["list-style-type"] = "disc";
                    break;

                case "decimal":
                    style["list-style-type"] = "decimal";
                    break;

                case "lowerLetter":
                    style["list-style-type"] = "lower-alpha";
                    break;

                case "upperLetter":
                    style["list-style-type"] = "upper-alpha";
                    break;

                case "lowerRoman":
                    style["list-style-type"] = "lower-roman";
                    break;

                case "upperRoman":
                    style["list-style-type"] = "upper-roman";
                    break;

                case "none":
                    style["list-style-type"] = "none";
                    break;
            }
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

	                case "pgSz":
                        if(!this.ignoreWidth)
                            elem.style["width"] = xml.sizeAttr(n, "w");

                        if(!this.ignoreHeight)
                            elem.style["height"] = xml.sizeAttr(n, "h");
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

                    case "hyperlink":
                        result.children.push(this.parseHyperlink(c));
                        break;

                    case "bookmarkStart":
                        result.children.push(this.parseBookmark(c));
                        break;

                    case "pPr":
                        this.parseParagraphProperties(c, result);
                        break;
                }
            });

            return result;
        }

        parseParagraphProperties(node: Node, paragraph: IDomParagraph) {
            this.parseDefaultProperties(node, paragraph.style = {}, null, c => {
                switch (c.localName) {
                    case "pStyle":
                        paragraph.className = xml.stringAttr(c, "val");
                        break;
                    
                    case "numPr":
                        this.parseNumbering(c, paragraph);
                        break;

                    case "framePr":
                        this.parseFrame(c, paragraph);
                        break;

                    case "rPr":
                        //TODO ignore
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

        parseFrame(node: Node, paragraph: IDomParagraph){
            var dropCap = xml.stringAttr(node, "dropCap");

            if(dropCap == "drop")
                paragraph.style["float"] = "left";
        }

        parseBookmark(node: Node): IDomElement {
            var result: IDomRun = { domType: DomType.Run };

            result.id = xml.stringAttr(node, "name");

            return result;
        }

        parseHyperlink(node: Node): IDomRun {
            var result: IDomHyperlink = { domType: DomType.Hyperlink, children: [] };
            var anchor = xml.stringAttr(node, "anchor");

            if(anchor)
                result.href = "#" + anchor;   

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "r":
                        result.children.push(this.parseRun(c));
                        break;
                }
            });     
            
            return result;
        }

        parseRun(node: Node): IDomRun {
            var result: IDomRun = { domType: DomType.Run };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "t":
                        result.text = c.textContent;//.replace(" ", "\u00A0"); // TODO
                        break;

                    case "br":
                        result.break = xml.stringAttr(c, "type") || "textWrapping";
                        break;

                    case "tab":
                        //result.text = "\u00A0\u00A0\u00A0\u00A0";  // TODO
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

	                case "tblGrid":
                        result.columns = this.parseTableColumns(c);
                        break;

                    case "tblPr":
                        this.parseTableProperties(c, result);
                        break;
                }
            });

            return result;
        }
        
        parseTableColumns(node: Node): IDomTableColumn[] {
            var result = [];
            
            xml.foreach(node, n => {
                switch (n.localName) {
                    case "gridCol":
                        result.push({ width: xml.sizeAttr(n, "w") });
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

        parseTableRow(node: Node): IDomTableRow {
            var result: IDomTableRow = { domType: DomType.Row, children: [] };

            xml.foreach(node, c => {
                switch (c.localName) {
                    case "tc":
                        result.children.push(this.parseTableCell(c));
                        break;

                    case "trPr":
                        this.parseTableRowProperties(c, result);
                        break;
                }
            });

            return result;
        }

        parseTableRowProperties(node: Node, row: IDomTableRow) {
            row.style = this.parseDefaultProperties(node, {}, null, c => {
                switch (c.localName) {
                    case "cnfStyle":
                        row.className = values.classNameOfCnfStyle(c);
                        break;

                    default:
                        return false;
                }

                return true;
            });
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

                    case "cnfStyle":
                        cell.className = values.classNameOfCnfStyle(c);
                        break;

                    default:
                        return false;
                }

                return true;
            });
        }

        parseDefaultProperties(node: Node, style: IDomStyleValues = null, childStyle: IDomStyleValues = null, handler: (prop: Node) => void = null): IDomStyleValues {
            style = style || {};

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
                        style["background-color"] = xml.colorAttr(c, "fill");
                        break;

                    case "highlight":
                        style["background-color"] = xml.colorAttr(c, "val");
                        break;

	                case "tcW": 
                        if(this.ignoreWidth)
                        break;

	                case "tblW":
                        style["width"] = xml.sizeAttr(c, "w");
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
                        this.parseIndentation(c, style);
                        break;

                    case "rFonts":
                        style["font-family"] = values.valueOfFonts(c);
                        break;

                    case "tblBorders":
                        this.parseBorderProperties(c, childStyle || style);
                        break;

                    case "tblCellSpacing":
                        style["border-spacing"] = values.valueOfMargin(c);
                        style["border-collapse"] = "separate";
                        break;

                    case "pBdr":
                        this.parseBorderProperties(c, style);
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

                    case "vAlign":
                        style["vertical-align"] = xml.stringAttr(c, "val");
                        break;

                    case "spacing":
                        this.parseSpacing(c, style);
                        break;

                    case "tabs":
                        this.parseTabs(c, style);
                        break;

                    case "lang":
                        //TODO ignore
                        break;

                    default:
                        if (handler != null && !handler(c))
                            this.debug && console.warn(`DOCX: Unknown document element: ${c.localName}`);
                        break;
                }
            });

            return style;
        }

        parseIndentation(node: Node, style: IDomStyleValues){
            var firstLine = xml.sizeAttr(node, "firstLine"); 
            var left = xml.sizeAttr(node, "left");
            var start = xml.sizeAttr(node, "start");
            var right = xml.sizeAttr(node, "right");
            var end = xml.sizeAttr(node, "end");

            if(firstLine) style["text-indent"] = firstLine;
            if(left || start) style["margin-left"] = left || start;
            if(right || end) style["margin-right"] = right || end;
        }

        parseSpacing(node: Node, style: IDomStyleValues) {
            var before = xml.sizeAttr(node, "before");
            var after = xml.sizeAttr(node, "after");
            var line = xml.sizeAttr(node, "line");

            if(before) style["margin-top"] = before;
            if(after) style["margin-bottom"] = after;
            if(line){ 
                style["line-height"] = line;
                style["min-height"] = line;
            }
        }

	    parseTabs(node: Node, style: IDomStyleValues) {
            xml.foreach(node, n => {
                switch(n.localName){
                    case "tab":{
                        var type = xml.stringAttr(n, "val");
                        var pos = xml.sizeAttr(n, "pos");

                        switch(type) {
                            case "left":
                                style["magrin-left"] = values.addSize(style["magrin-left"], pos);
                                break;

                            case "right":
                                style["magrin-right"] = values.addSize(style["magrin-right"], pos);
                                break;
                        }
                    }
                        break;                    
                }
            });
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
                    case "start":
                    case "left":
                        output["border-left"] = values.valueOfBorder(c);
                        break;

                    case "end":
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

        static nodes(node: Node) {
            var result = [];
            for (var i = 0; i < node.childNodes.length; i++)
                result.push(node.childNodes[i]);
            return result;
        }

        static foreach(node: Node, cb: (n: Node) => void) {
            for (var i = 0; i < node.childNodes.length; i++)
                cb(node.childNodes[i]);
        }

        static byTagName(node: Node, tagName: string) {
            for (var i = 0; i < node.childNodes.length; i++)
                if (node.childNodes[i].localName == tagName)
                    return node.childNodes[i];
        }

        static nodeStringAttr(node: Node, nodeName, attrName: string) {
            var n = xml.byTagName(node, nodeName)
            return n ? xml.stringAttr(n, attrName) : null;
        }

        static stringAttr(node: Node, attrName: string) {
            for (var i = 0; i < node.attributes.length; i++) {
                var attr = node.attributes.item(i);

                if (attr.localName == attrName)
                    return attr.value;
            }

            return null;
        }

        static colorAttr(node: Node, attrName: string, defValue: string = null) {
            var v = xml.stringAttr(node, attrName);
            
            switch (v)
            {
                case "yellow":
                     return v;

                case "auto":
                     return "black"
            }

            return v ? `#${v}` : defValue;
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
                case SizeType.Dxa: return (0.05 * intVal).toFixed(2) + "pt";
                case SizeType.FontSize: return (0.5 * intVal).toFixed(2) + "pt";
                case SizeType.Percent: return (0.01 * intVal).toFixed(2) + "%";
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

        static valueOfRelType(c: Node) {
            switch(xml.sizeAttr(c, "Type")) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings": 
                    return DomRelationshipType.Settings;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                    return DomRelationshipType.Theme;
                case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects": 
                    return DomRelationshipType.StylesWithEffects;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                    return DomRelationshipType.Styles;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable": 
                    return DomRelationshipType.FontTable;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image": 
                    return DomRelationshipType.Image;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings": 
                    return DomRelationshipType.WebSettings;
            }

            return DomRelationshipType.Unknown;
        }

        static valueOfBorder(c: Node) {
            var type = xml.stringAttr(c, "val");

            if (type == "nil")
                return "none";

            var color = xml.colorAttr(c, "color");
            var size = xml.sizeAttr(c, "sz");

            return `${size} solid ${color == "auto" ? "black" : color}`;
        }

        static valueOfTblLayout(c: Node) {
            var type = xml.stringAttr(c, "val");
            return type == "fixed" ? "fixed" : "auto";
        }

        static classNameOfCnfStyle(c: Node){
            let className = "";
            let val = xml.stringAttr(c, "val");
            //FirstRow, LastRow, FirstColumn, LastColumn, Band1Vertical, Band2Vertical, Band1Horizontal, Band2Horizontal, NE Cell, NW Cell, SE Cell, SW Cell.

            if(val[0] == "1") className += " first-row";
            if(val[1] == "1") className += " last-row";
            if(val[2] == "1") className += " first-col";
            if(val[3] == "1") className += " last-col";
            if(val[4] == "1") className += " odd-col";
            if(val[5] == "1") className += " even-col";
            if(val[6] == "1") className += " odd-row";
            if(val[7] == "1") className += " even-row";
            if(val[8] == "1") className += " ne-cell";
            if(val[9] == "1") className += " nw-cell";
            if(val[10] == "1") className += " se-cell";
            if(val[11] == "1") className += " sw-cell";
            
            return className.trim();
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

        static valueOfFonts(c: Node){
            var ascii = xml.stringAttr(c, "ascii");
            return ascii;
        }

        static addSize(a: string, b: string): string {
            if(a == null) return b;
            if(b == null) return a;

            return `calc(${a} + ${b})`; //TODO
        }
    }
}
