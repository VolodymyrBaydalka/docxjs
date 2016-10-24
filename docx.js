var docx;
(function (docx) {
    var DocumentParser = (function () {
        function DocumentParser() {
            this.skipDeclaration = true;
        }
        DocumentParser.prototype.parseDocumentAsync = function (zip) {
            var _this = this;
            return zip.files["word/document.xml"]
                .async("string")
                .then(function (xml) { return _this.parseDocumentFile(xml); });
        };
        DocumentParser.prototype.parseStylesAsync = function (zip) {
            var _this = this;
            return zip.files["word/styles.xml"]
                .async("string")
                .then(function (xml) { return _this.parseStylesFile(xml); });
        };
        DocumentParser.prototype.parseDocumentFile = function (xmlString) {
            var result = {
                domType: docx.DomType.Document,
                children: [],
                style: {}
            };
            var xbody = xml.byTagName(xml.parse(xmlString, this.skipDeclaration), "body");
            for (var i = 0; i < xbody.childNodes.length; i++) {
                var node = xbody.childNodes[i];
                switch (node.localName) {
                    case "p":
                        result.children.push(parseParagraph(node));
                        break;
                    case "tbl":
                        result.children.push(parseTable(node));
                        break;
                    case "sectPr":
                        this.parseSectionProperties(node, result);
                        break;
                }
            }
            return result;
        };
        DocumentParser.prototype.parseStylesFile = function (xmlString) {
            var _this = this;
            var result = [];
            var xstyles = xml.parse(xmlString, this.skipDeclaration);
            xml.foreach(xstyles, function (n) {
                switch (n.localName) {
                    case "style":
                        result.push(_this.parseStyle(n));
                        break;
                }
            });
            return result;
        };
        DocumentParser.prototype.parseStyle = function (node) {
            var result = {
                id: xml.stringAttr(node, "styleId"),
                target: null
            };
            switch (xml.stringAttr(node, "type")) {
                case "paragraph":
                    result.target = "p";
                    break;
                case "table":
                    result.target = "table";
                    break;
            }
            return result;
        };
        DocumentParser.prototype.parseSectionProperties = function (node, elem) {
            xml.foreach(node, function (n) {
                switch (n.localName) {
                    case "pgMar":
                        elem.style["padding-left"] = xml.sizeAttr(n, "left");
                        elem.style["padding-right"] = xml.sizeAttr(n, "right");
                        elem.style["padding-top"] = xml.sizeAttr(n, "top");
                        elem.style["padding-bottom"] = xml.sizeAttr(n, "bottom");
                        break;
                }
            });
        };
        return DocumentParser;
    }());
    docx.DocumentParser = DocumentParser;
    function parseParagraph(node) {
        var result = { domType: docx.DomType.Paragraph, children: [] };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "r":
                    result.children.push(parseRun(c));
                    break;
                case "pPr":
                    result.style = parseDefaultProperties(c, {}, null);
                    break;
            }
        });
        return result;
    }
    function parseRun(node) {
        var result = { domType: docx.DomType.Run };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "t":
                    result.text = c.textContent;
                    break;
                case "br":
                    result.isBreak = true;
                    break;
                case "rPr":
                    result.style = parseDefaultProperties(c, {}, null);
                    break;
            }
        });
        return result;
    }
    function parseTable(node) {
        var result = { domType: docx.DomType.Table, children: [] };
        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];
            switch (c.localName) {
                case "tr":
                    result.children.push(parseTableRow(c));
                    break;
            }
        }
        return result;
    }
    function parseTableRow(node) {
        var result = { domType: docx.DomType.Row, children: [] };
        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];
            switch (c.localName) {
                case "tc":
                    result.children.push(parseTableCell(c));
                    break;
            }
        }
        return result;
    }
    function parseTableCell(node) {
        var result = { domType: docx.DomType.Cell, children: [] };
        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];
            switch (c.localName) {
                case "tbl":
                    result.children.push(parseTable(c));
                    break;
                case "p":
                    result.children.push(parseParagraph(c));
                    break;
            }
        }
        return result;
    }
    function parseDefaultProperties(node, output, handler) {
        xml.foreach(node, function (c) {
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
    var SizeType;
    (function (SizeType) {
        SizeType[SizeType["FontSize"] = 0] = "FontSize";
        SizeType[SizeType["Dxa"] = 1] = "Dxa";
        SizeType[SizeType["Percent"] = 2] = "Percent";
    })(SizeType || (SizeType = {}));
    var xml = (function () {
        function xml() {
        }
        xml.parse = function (xmlString, skipDeclaration) {
            if (skipDeclaration === void 0) { skipDeclaration = true; }
            if (skipDeclaration)
                xmlString = xmlString.replace(/<[?].*[?]>/, "");
            return new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
        };
        xml.foreach = function (node, cb) {
            for (var i = 0; i < node.childNodes.length; i++)
                cb(node.childNodes.item(i));
        };
        xml.byTagName = function (node, tagName) {
            for (var i = 0; i < node.childNodes.length; i++)
                if (node.childNodes[i].localName == tagName)
                    return node.childNodes[i];
        };
        xml.stringAttr = function (node, attrName) {
            for (var i = 0; i < node.attributes.length; i++) {
                var attr = node.attributes.item(i);
                if (attr.localName == attrName)
                    return attr.value;
            }
            return null;
        };
        xml.sizeAttr = function (node, attrName, type) {
            if (type === void 0) { type = SizeType.Dxa; }
            var val = xml.stringAttr(node, attrName);
            var intVal = parseInt(val);
            switch (type) {
                case SizeType.Dxa: return (0.05 * intVal) + "pt";
                case SizeType.FontSize: return (0.5 * intVal) + "pt";
                case SizeType.Percent: return (0.01 * intVal) + "%";
            }
            return val;
        };
        return xml;
    }());
})(docx || (docx = {}));
var docx;
(function (docx) {
    (function (DomType) {
        DomType[DomType["Document"] = 0] = "Document";
        DomType[DomType["Paragraph"] = 1] = "Paragraph";
        DomType[DomType["Run"] = 2] = "Run";
        DomType[DomType["Break"] = 3] = "Break";
        DomType[DomType["Table"] = 4] = "Table";
        DomType[DomType["Row"] = 5] = "Row";
        DomType[DomType["Cell"] = 6] = "Cell";
    })(docx.DomType || (docx.DomType = {}));
    var DomType = docx.DomType;
})(docx || (docx = {}));
var docx;
(function (docx) {
    var HtmlRenderer = (function () {
        function HtmlRenderer(htmlDocument) {
            this.htmlDocument = htmlDocument;
        }
        HtmlRenderer.prototype.renderDocument = function (document) {
            var bodyElement = this.htmlDocument.createElement("section");
            this.renderChildren(document, bodyElement);
            this.renderStyleValues(document, bodyElement);
            return bodyElement;
        };
        HtmlRenderer.prototype.renderStyles = function (styles) {
            var styleElement = document.createElement("style");
            var styleText = "";
            styleElement.type = "text/css";
            for (var _i = 0, styles_1 = styles; _i < styles_1.length; _i++) {
                var style = styles_1[_i];
                if (style.isDefault)
                    styleText += style.target + ", ";
                styleText += style.target + "." + style.id + "{\r\n";
                styleText += "}\r\n";
            }
            styleElement.innerHTML = styleText;
            return styleElement;
        };
        HtmlRenderer.prototype.renderElement = function (elem) {
            switch (elem.domType) {
                case docx.DomType.Paragraph:
                    return this.renderParagraph(elem);
                case docx.DomType.Run:
                    return this.renderRun(elem);
                case docx.DomType.Table:
                    return this.renderTable(elem);
                case docx.DomType.Row:
                    return this.renderTableRow(elem);
                case docx.DomType.Cell:
                    return this.renderTableCell(elem);
            }
            return null;
        };
        HtmlRenderer.prototype.renderChildren = function (elem, into) {
            var _this = this;
            var result = null;
            if (elem.children != null)
                result = elem.children.map(function (x) { return _this.renderElement(x); }).filter(function (x) { return x != null; });
            if (into && result)
                result.forEach(function (x) { return into.appendChild(x); });
            return result;
        };
        HtmlRenderer.prototype.renderParagraph = function (elem) {
            var result = this.htmlDocument.createElement("p");
            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);
            return result;
        };
        HtmlRenderer.prototype.renderRun = function (elem) {
            if (elem.isBreak)
                return this.htmlDocument.createElement("br");
            var result = this.htmlDocument.createElement("span");
            this.renderStyleValues(elem, result);
            result.textContent = elem.text;
            return result;
        };
        HtmlRenderer.prototype.renderTable = function (elem) {
            var result = this.htmlDocument.createElement("table");
            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);
            return result;
        };
        HtmlRenderer.prototype.renderTableRow = function (elem) {
            var result = this.htmlDocument.createElement("tr");
            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);
            return result;
        };
        HtmlRenderer.prototype.renderTableCell = function (elem) {
            var result = this.htmlDocument.createElement("td");
            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);
            return result;
        };
        HtmlRenderer.prototype.renderStyleValues = function (input, ouput) {
            if (input.style == null)
                return;
            for (var key in input.style) {
                if (input.style.hasOwnProperty(key)) {
                    ouput.style[key] = input.style[key];
                }
            }
        };
        return HtmlRenderer;
    }());
    docx.HtmlRenderer = HtmlRenderer;
})(docx || (docx = {}));
var docx;
(function (docx) {
    function renderAsync(data, bodyContainer, styleContainer) {
        if (styleContainer === void 0) { styleContainer = null; }
        var parser = new docx.DocumentParser();
        var renderer = new docx.HtmlRenderer(window.document);
        var _zip = null;
        var _doc = null;
        return JSZip.loadAsync(data)
            .then(function (zip) { _zip = zip; return parser.parseDocumentAsync(_zip); })
            .then(function (doc) { _doc = doc; return parser.parseStylesAsync(_zip); })
            .then(function (styles) {
            styleContainer = styleContainer || bodyContainer;
            clearElement(styleContainer);
            clearElement(bodyContainer);
            styleContainer.appendChild(renderer.renderStyles(styles));
            bodyContainer.appendChild(renderer.renderDocument(_doc));
            return { document: _doc, styles: styles };
        });
    }
    docx.renderAsync = renderAsync;
    function clearElement(elem) {
        while (elem.firstChild) {
            elem.removeChild(elem.firstChild);
        }
    }
})(docx || (docx = {}));
//# sourceMappingURL=docx.js.map