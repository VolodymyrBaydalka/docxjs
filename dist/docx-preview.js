(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory(require("JSZip"));
	else if(typeof define === 'function' && define.amd)
		define(["JSZip"], factory);
	else if(typeof exports === 'object')
		exports["docx"] = factory(require("JSZip"));
	else
		root["docx"] = factory(root["JSZip"]);
})(self, function(__WEBPACK_EXTERNAL_MODULE_jszip__) {
return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/common/open-xml-package.ts":
/*!****************************************!*\
  !*** ./src/common/open-xml-package.ts ***!
  \****************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.OpenXmlPackage = void 0;
var JSZip = __webpack_require__(/*! jszip */ "jszip");
var xml_parser_1 = __webpack_require__(/*! ../parser/xml-parser */ "./src/parser/xml-parser.ts");
var utils_1 = __webpack_require__(/*! ../utils */ "./src/utils.ts");
var relationship_1 = __webpack_require__(/*! ./relationship */ "./src/common/relationship.ts");
var OpenXmlPackage = (function () {
    function OpenXmlPackage(_zip, options) {
        this._zip = _zip;
        this.options = options;
        this.xmlParser = new xml_parser_1.XmlParser();
    }
    OpenXmlPackage.prototype.get = function (path) {
        return this._zip.files[normalizePath(path)];
    };
    OpenXmlPackage.prototype.update = function (path, content) {
        this._zip.file(path, content);
    };
    OpenXmlPackage.load = function (input, options) {
        return JSZip.loadAsync(input).then(function (zip) { return new OpenXmlPackage(zip, options); });
    };
    OpenXmlPackage.prototype.save = function (type) {
        if (type === void 0) { type = "blob"; }
        return this._zip.generateAsync({ type: type });
    };
    OpenXmlPackage.prototype.load = function (path, type) {
        var _a, _b;
        if (type === void 0) { type = "string"; }
        return (_b = (_a = this.get(path)) === null || _a === void 0 ? void 0 : _a.async(type)) !== null && _b !== void 0 ? _b : Promise.resolve(null);
    };
    OpenXmlPackage.prototype.loadRelationships = function (path) {
        var _this = this;
        if (path === void 0) { path = null; }
        var relsPath = "_rels/.rels";
        if (path != null) {
            var _a = (0, utils_1.splitPath)(path), f = _a[0], fn = _a[1];
            relsPath = "".concat(f, "_rels/").concat(fn, ".rels");
        }
        return this.load(relsPath)
            .then(function (txt) { return txt ? (0, relationship_1.parseRelationships)(_this.parseXmlDocument(txt).firstElementChild, _this.xmlParser) : null; });
    };
    OpenXmlPackage.prototype.parseXmlDocument = function (txt) {
        return (0, xml_parser_1.parseXmlString)(txt, this.options.trimXmlDeclaration);
    };
    return OpenXmlPackage;
}());
exports.OpenXmlPackage = OpenXmlPackage;
function normalizePath(path) {
    return path.startsWith('/') ? path.substr(1) : path;
}


/***/ }),

/***/ "./src/common/part.ts":
/*!****************************!*\
  !*** ./src/common/part.ts ***!
  \****************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.Part = void 0;
var xml_parser_1 = __webpack_require__(/*! ../parser/xml-parser */ "./src/parser/xml-parser.ts");
var Part = (function () {
    function Part(_package, path) {
        this._package = _package;
        this.path = path;
    }
    Part.prototype.load = function () {
        var _this = this;
        return Promise.all([
            this._package.loadRelationships(this.path).then(function (rels) {
                _this.rels = rels;
            }),
            this._package.load(this.path).then(function (text) {
                var xmlDoc = _this._package.parseXmlDocument(text);
                if (_this._package.options.keepOrigin) {
                    _this._xmlDocument = xmlDoc;
                }
                _this.parseXml(xmlDoc.firstElementChild);
            })
        ]);
    };
    Part.prototype.save = function () {
        this._package.update(this.path, (0, xml_parser_1.serializeXmlString)(this._xmlDocument));
    };
    Part.prototype.parseXml = function (root) {
    };
    return Part;
}());
exports.Part = Part;


/***/ }),

/***/ "./src/common/relationship.ts":
/*!************************************!*\
  !*** ./src/common/relationship.ts ***!
  \************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseRelationships = exports.RelationshipTypes = void 0;
var RelationshipTypes;
(function (RelationshipTypes) {
    RelationshipTypes["OfficeDocument"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    RelationshipTypes["FontTable"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    RelationshipTypes["Image"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    RelationshipTypes["Numbering"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    RelationshipTypes["Styles"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    RelationshipTypes["StylesWithEffects"] = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
    RelationshipTypes["Theme"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    RelationshipTypes["Settings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    RelationshipTypes["WebSettings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
    RelationshipTypes["Hyperlink"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    RelationshipTypes["Footnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    RelationshipTypes["Footer"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    RelationshipTypes["Header"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    RelationshipTypes["ExtendedProperties"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    RelationshipTypes["CoreProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
})(RelationshipTypes = exports.RelationshipTypes || (exports.RelationshipTypes = {}));
function parseRelationships(root, xml) {
    return xml.elements(root).map(function (e) { return ({
        id: xml.attr(e, "Id"),
        type: xml.attr(e, "Type"),
        target: xml.attr(e, "Target"),
        targetMode: xml.attr(e, "TargetMode")
    }); });
}
exports.parseRelationships = parseRelationships;


/***/ }),

/***/ "./src/document-parser.ts":
/*!********************************!*\
  !*** ./src/document-parser.ts ***!
  \********************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.DocumentParser = exports.autos = void 0;
var dom_1 = __webpack_require__(/*! ./document/dom */ "./src/document/dom.ts");
var utils = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var paragraph_1 = __webpack_require__(/*! ./document/paragraph */ "./src/document/paragraph.ts");
var section_1 = __webpack_require__(/*! ./document/section */ "./src/document/section.ts");
var xml_parser_1 = __webpack_require__(/*! ./parser/xml-parser */ "./src/parser/xml-parser.ts");
var bookmark_1 = __webpack_require__(/*! ./document/bookmark */ "./src/document/bookmark.ts");
var footer_1 = __webpack_require__(/*! ./footer/footer */ "./src/footer/footer.ts");
var header_1 = __webpack_require__(/*! ./header/header */ "./src/header/header.ts");
var footnote_1 = __webpack_require__(/*! ./footnotes/footnote */ "./src/footnotes/footnote.ts");
exports.autos = {
    shd: "white",
    color: "black",
    highlight: "transparent"
};
var DocumentParser = (function () {
    function DocumentParser(options) {
        this.options = __assign({ ignoreWidth: false, debug: false }, options);
    }
    DocumentParser.prototype.parseFootnotes = function (xmlDoc) {
        var result = [];
        for (var _i = 0, _a = xml_parser_1.default.elements(xmlDoc, "footnote"); _i < _a.length; _i++) {
            var el = _a[_i];
            var footnote = new footnote_1.WmlFootnote();
            footnote.id = xml_parser_1.default.attr(el, "id");
            footnote.footnoteType = xml_parser_1.default.attr(el, "type");
            footnote.children = this.parseBodyElements(el);
            result.push(footnote);
        }
        return result;
    };
    DocumentParser.prototype.parseFooter = function (xmlDoc) {
        var result = new footer_1.WmlFooter();
        result.children = this.parseBodyElements(xmlDoc);
        return result;
    };
    DocumentParser.prototype.parseHeader = function (xmlDoc) {
        var result = new header_1.WmlHeader();
        result.children = this.parseBodyElements(xmlDoc);
        return result;
    };
    DocumentParser.prototype.parseDocumentFile = function (xmlDoc) {
        var xbody = xml_parser_1.default.element(xmlDoc, "body");
        var background = xml_parser_1.default.element(xmlDoc, "background");
        var sectPr = xml_parser_1.default.element(xbody, "sectPr");
        return {
            type: dom_1.DomType.Document,
            children: this.parseBodyElements(xbody),
            props: sectPr ? (0, section_1.parseSectionProperties)(sectPr, xml_parser_1.default) : null,
            cssStyle: background ? this.parseBackground(background) : {},
        };
    };
    DocumentParser.prototype.parseBackground = function (elem) {
        var result = {};
        var color = xml.colorAttr(elem, "color");
        if (color) {
            result["background-color"] = color;
        }
        return result;
    };
    DocumentParser.prototype.parseBodyElements = function (element) {
        var _this = this;
        var children = [];
        xml.foreach(element, function (elem) {
            switch (elem.localName) {
                case "p":
                    children.push(_this.parseParagraph(elem));
                    break;
                case "tbl":
                    children.push(_this.parseTable(elem));
                    break;
            }
        });
        return children;
    };
    DocumentParser.prototype.parseStylesFile = function (xstyles) {
        var _this = this;
        var result = [];
        xml.foreach(xstyles, function (n) {
            switch (n.localName) {
                case "style":
                    result.push(_this.parseStyle(n));
                    break;
                case "docDefaults":
                    result.push(_this.parseDefaultStyles(n));
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseDefaultStyles = function (node) {
        var _this = this;
        var result = {
            id: null,
            name: null,
            target: null,
            basedOn: null,
            styles: []
        };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "rPrDefault":
                    var rPr = xml_parser_1.default.element(c, "rPr");
                    if (rPr)
                        result.styles.push({
                            target: "span",
                            values: _this.parseDefaultProperties(rPr, {})
                        });
                    break;
                case "pPrDefault":
                    var pPr = xml_parser_1.default.element(c, "pPr");
                    if (pPr)
                        result.styles.push({
                            target: "p",
                            values: _this.parseDefaultProperties(pPr, {})
                        });
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseStyle = function (node) {
        var _this = this;
        var result = {
            id: xml.stringAttr(node, "styleId"),
            isDefault: xml.boolAttr(node, "default"),
            name: null,
            target: null,
            basedOn: null,
            styles: [],
            linked: null
        };
        switch (xml.stringAttr(node, "type")) {
            case "paragraph":
                result.target = "p";
                break;
            case "table":
                result.target = "table";
                break;
            case "character":
                result.target = "span";
                break;
        }
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "basedOn":
                    result.basedOn = xml.className(n, "val");
                    break;
                case "name":
                    result.name = xml.stringAttr(n, "val");
                    break;
                case "link":
                    result.linked = xml.className(n, "val");
                    break;
                case "next":
                    result.next = xml.className(n, "val");
                    break;
                case "aliases":
                    result.aliases = xml.stringAttr(n, "val").split(",");
                    break;
                case "pPr":
                    result.styles.push({
                        target: "p",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    result.paragraphProps = (0, paragraph_1.parseParagraphProperties)(n, xml_parser_1.default);
                    break;
                case "rPr":
                    result.styles.push({
                        target: "span",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    result.runProps = (0, paragraph_1.parseParagraphProperties)(n, xml_parser_1.default);
                    break;
                case "tblPr":
                case "tcPr":
                    result.styles.push({
                        target: "td",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "tblStylePr":
                    for (var _i = 0, _a = _this.parseTableStyle(n); _i < _a.length; _i++) {
                        var s = _a[_i];
                        result.styles.push(s);
                    }
                    break;
                case "rsid":
                case "qFormat":
                case "hidden":
                case "semiHidden":
                case "unhideWhenUsed":
                case "autoRedefine":
                case "uiPriority":
                    break;
                default:
                    _this.options.debug && console.warn("DOCX: Unknown style element: ".concat(n.localName));
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableStyle = function (node) {
        var _this = this;
        var result = [];
        var type = xml.stringAttr(node, "type");
        var selector = "";
        switch (type) {
            case "firstRow":
                selector = "tr.first-row td";
                break;
            case "lastRow":
                selector = "tr.last-row td";
                break;
            case "firstCol":
                selector = "td.first-col";
                break;
            case "lastCol":
                selector = "td.last-col";
                break;
            case "band1Vert":
                selector = "td.odd-col";
                break;
            case "band2Vert":
                selector = "td.even-col";
                break;
            case "band1Horz":
                selector = "tr.odd-row";
                break;
            case "band2Horz":
                selector = "tr.even-row";
                break;
            default: return [];
        }
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "pPr":
                    result.push({
                        target: selector + " p",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "rPr":
                    result.push({
                        target: selector + " span",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "tblPr":
                case "tcPr":
                    result.push({
                        target: selector,
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseNumberingFile = function (xnums) {
        var _this = this;
        var result = [];
        var mapping = {};
        var bullets = [];
        xml.foreach(xnums, function (n) {
            switch (n.localName) {
                case "abstractNum":
                    _this.parseAbstractNumbering(n, bullets)
                        .forEach(function (x) { return result.push(x); });
                    break;
                case "numPicBullet":
                    bullets.push(_this.parseNumberingPicBullet(n));
                    break;
                case "num":
                    var numId = xml.stringAttr(n, "numId");
                    var abstractNumId = xml.elementStringAttr(n, "abstractNumId", "val");
                    mapping[abstractNumId] = numId;
                    break;
            }
        });
        result.forEach(function (x) { return x.id = mapping[x.id]; });
        return result;
    };
    DocumentParser.prototype.parseNumberingPicBullet = function (elem) {
        var pict = xml_parser_1.default.element(elem, "pict");
        var shape = pict && xml_parser_1.default.element(pict, "shape");
        var imagedata = shape && xml_parser_1.default.element(shape, "imagedata");
        return imagedata ? {
            id: xml.intAttr(elem, "numPicBulletId"),
            src: xml.stringAttr(imagedata, "id"),
            style: xml.stringAttr(shape, "style")
        } : null;
    };
    DocumentParser.prototype.parseAbstractNumbering = function (node, bullets) {
        var _this = this;
        var result = [];
        var id = xml.stringAttr(node, "abstractNumId");
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "lvl":
                    result.push(_this.parseNumberingLevel(id, n, bullets));
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseNumberingLevel = function (id, node, bullets) {
        var _this = this;
        var result = {
            id: id,
            level: xml.intAttr(node, "ilvl"),
            pStyle: {},
            rStyle: {},
            suff: "tab"
        };
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "pPr":
                    _this.parseDefaultProperties(n, result.pStyle);
                    break;
                case "rPr":
                    _this.parseDefaultProperties(n, result.rStyle);
                    break;
                case "lvlPicBulletId":
                    var id = xml.intAttr(n, "val");
                    result.bullet = bullets.filter(function (x) { return x.id == id; })[0];
                    break;
                case "lvlText":
                    result.levelText = xml.stringAttr(n, "val");
                    break;
                case "numFmt":
                    result.format = xml.stringAttr(n, "val");
                    break;
                case "suff":
                    result.suff = xml.stringAttr(n, "val");
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseParagraph = function (node) {
        var _this = this;
        var result = { type: dom_1.DomType.Paragraph, children: [] };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "r":
                    result.children.push(_this.parseRun(c, result));
                    break;
                case "hyperlink":
                    result.children.push(_this.parseHyperlink(c, result));
                    break;
                case "bookmarkStart":
                    result.children.push((0, bookmark_1.parseBookmarkStart)(c, xml_parser_1.default));
                    break;
                case "bookmarkEnd":
                    result.children.push((0, bookmark_1.parseBookmarkEnd)(c, xml_parser_1.default));
                    break;
                case "pPr":
                    _this.parseParagraphProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseParagraphProperties = function (elem, paragraph) {
        var _this = this;
        this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, function (c) {
            if ((0, paragraph_1.parseParagraphProperty)(c, paragraph, xml_parser_1.default))
                return true;
            switch (c.localName) {
                case "pStyle":
                    utils.addElementClass(paragraph, xml.className(c, "val"));
                    break;
                case "cnfStyle":
                    utils.addElementClass(paragraph, values.classNameOfCnfStyle(c));
                    break;
                case "framePr":
                    _this.parseFrame(c, paragraph);
                    break;
                case "rPr":
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseFrame = function (node, paragraph) {
        var dropCap = xml.stringAttr(node, "dropCap");
        if (dropCap == "drop")
            paragraph.cssStyle["float"] = "left";
    };
    DocumentParser.prototype.parseHyperlink = function (node, parent) {
        var _this = this;
        var result = { type: dom_1.DomType.Hyperlink, parent: parent, children: [] };
        var anchor = xml.stringAttr(node, "anchor");
        if (anchor)
            result.href = "#" + anchor;
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "r":
                    result.children.push(_this.parseRun(c, result));
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseRun = function (node, parent) {
        var _this = this;
        var result = { type: dom_1.DomType.Run, parent: parent, children: [] };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "t":
                    result.children.push({
                        type: dom_1.DomType.Text,
                        text: c.textContent
                    });
                    break;
                case "fldChar":
                    result.fldCharType = xml.stringAttr(c, "fldCharType");
                    break;
                case "br":
                    result.children.push({
                        type: dom_1.DomType.Break,
                        break: xml.stringAttr(c, "type") || "textWrapping"
                    });
                    break;
                case "lastRenderedPageBreak":
                    result.children.push({
                        type: dom_1.DomType.Break,
                        break: "lastRenderedPageBreak"
                    });
                    break;
                case "sym":
                    result.children.push({
                        type: dom_1.DomType.Symbol,
                        font: xml.stringAttr(c, "font"),
                        char: xml.stringAttr(c, "char")
                    });
                    break;
                case "tab":
                    result.children.push({ type: dom_1.DomType.Tab });
                    break;
                case "footnoteReference":
                    result.children.push({
                        type: dom_1.DomType.FootnoteReference,
                        id: xml.stringAttr(c, "id")
                    });
                    break;
                case "instrText":
                    result.instrText = c.textContent;
                    break;
                case "drawing":
                    var d = _this.parseDrawing(c);
                    if (d)
                        result.children = [d];
                    break;
                case "rPr":
                    _this.parseRunProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseRunProperties = function (elem, run) {
        this.parseDefaultProperties(elem, run.cssStyle = {}, null, function (c) {
            switch (c.localName) {
                case "rStyle":
                    run.className = xml.className(c, "val");
                    break;
                case "vertAlign":
                    switch (xml.stringAttr(c, "val")) {
                        case "subscript":
                            run.wrapper = "sub";
                            break;
                        case "superscript":
                            run.wrapper = "sup";
                            break;
                    }
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseDrawing = function (node) {
        for (var _i = 0, _a = xml_parser_1.default.elements(node); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "inline":
                case "anchor":
                    return this.parseDrawingWrapper(n);
            }
        }
    };
    DocumentParser.prototype.parseDrawingWrapper = function (node) {
        var result = { type: dom_1.DomType.Drawing, children: [], cssStyle: {} };
        var isAnchor = node.localName == "anchor";
        var wrapType = null;
        var simplePos = xml.boolAttr(node, "simplePos");
        var posX = { relative: "page", align: "left", offset: "0" };
        var posY = { relative: "page", align: "top", offset: "0" };
        for (var _i = 0, _a = xml_parser_1.default.elements(node); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "simplePos":
                    if (simplePos) {
                        posX.offset = xml.sizeAttr(n, "x", SizeType.Emu);
                        posY.offset = xml.sizeAttr(n, "y", SizeType.Emu);
                    }
                    break;
                case "extent":
                    result.cssStyle["width"] = xml.sizeAttr(n, "cx", SizeType.Emu);
                    result.cssStyle["height"] = xml.sizeAttr(n, "cy", SizeType.Emu);
                    break;
                case "positionH":
                case "positionV":
                    if (!simplePos) {
                        var pos = n.localName == "positionH" ? posX : posY;
                        var alignNode = xml_parser_1.default.element(n, "align");
                        var offsetNode = xml_parser_1.default.element(n, "posOffset");
                        if (alignNode)
                            pos.align = alignNode.textContent;
                        if (offsetNode)
                            pos.offset = xml.sizeValue(offsetNode, SizeType.Emu);
                    }
                    break;
                case "wrapTopAndBottom":
                    wrapType = "wrapTopAndBottom";
                    break;
                case "wrapNone":
                    wrapType = "wrapNone";
                    break;
                case "graphic":
                    var g = this.parseGraphic(n);
                    if (g)
                        result.children.push(g);
                    break;
            }
        }
        if (wrapType == "wrapTopAndBottom") {
            result.cssStyle['display'] = 'block';
            if (posX.align) {
                result.cssStyle['text-align'] = posX.align;
                result.cssStyle['width'] = "100%";
            }
        }
        else if (wrapType == "wrapNone") {
            result.cssStyle['display'] = 'block';
            result.cssStyle['position'] = 'relative';
            result.cssStyle["width"] = "0px";
            result.cssStyle["height"] = "0px";
            if (posX.offset)
                result.cssStyle["left"] = posX.offset;
            if (posY.offset)
                result.cssStyle["top"] = posY.offset;
        }
        else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
            result.cssStyle["float"] = posX.align;
        }
        return result;
    };
    DocumentParser.prototype.parseGraphic = function (elem) {
        var graphicData = xml_parser_1.default.element(elem, "graphicData");
        for (var _i = 0, _a = xml_parser_1.default.elements(graphicData); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "pic":
                    return this.parsePicture(n);
            }
        }
        return null;
    };
    DocumentParser.prototype.parsePicture = function (elem) {
        var result = { type: dom_1.DomType.Image, src: "", cssStyle: {} };
        var blipFill = xml_parser_1.default.element(elem, "blipFill");
        var blip = xml_parser_1.default.element(blipFill, "blip");
        result.src = xml.stringAttr(blip, "embed");
        var spPr = xml_parser_1.default.element(elem, "spPr");
        var xfrm = xml_parser_1.default.element(spPr, "xfrm");
        result.cssStyle["position"] = "relative";
        for (var _i = 0, _a = xml_parser_1.default.elements(xfrm); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "ext":
                    result.cssStyle["width"] = xml.sizeAttr(n, "cx", SizeType.Emu);
                    result.cssStyle["height"] = xml.sizeAttr(n, "cy", SizeType.Emu);
                    break;
                case "off":
                    result.cssStyle["left"] = xml.sizeAttr(n, "x", SizeType.Emu);
                    result.cssStyle["top"] = xml.sizeAttr(n, "y", SizeType.Emu);
                    break;
            }
        }
        return result;
    };
    DocumentParser.prototype.parseTable = function (node) {
        var _this = this;
        var result = { type: dom_1.DomType.Table, children: [] };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "tr":
                    result.children.push(_this.parseTableRow(c));
                    break;
                case "tblGrid":
                    result.columns = _this.parseTableColumns(c);
                    break;
                case "tblPr":
                    _this.parseTableProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableColumns = function (node) {
        var result = [];
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "gridCol":
                    result.push({ width: xml.sizeAttr(n, "w") });
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableProperties = function (elem, table) {
        var _this = this;
        table.cssStyle = {};
        table.cellStyle = {};
        this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, function (c) {
            switch (c.localName) {
                case "tblStyle":
                    table.className = xml.className(c, "val");
                    break;
                case "tblLook":
                    utils.addElementClass(table, values.classNameOftblLook(c));
                    break;
                case "tblpPr":
                    _this.parseTablePosition(c, table);
                    break;
                default:
                    return false;
            }
            return true;
        });
        switch (table.cssStyle["text-align"]) {
            case "center":
                delete table.cssStyle["text-align"];
                table.cssStyle["margin-left"] = "auto";
                table.cssStyle["margin-right"] = "auto";
                break;
            case "right":
                delete table.cssStyle["text-align"];
                table.cssStyle["margin-left"] = "auto";
                break;
        }
    };
    DocumentParser.prototype.parseTablePosition = function (node, table) {
        var topFromText = xml.sizeAttr(node, "topFromText");
        var bottomFromText = xml.sizeAttr(node, "bottomFromText");
        var rightFromText = xml.sizeAttr(node, "rightFromText");
        var leftFromText = xml.sizeAttr(node, "leftFromText");
        table.cssStyle["float"] = 'left';
        table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
        table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
        table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
        table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
    };
    DocumentParser.prototype.parseTableRow = function (node) {
        var _this = this;
        var result = { type: dom_1.DomType.Row, children: [] };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "tc":
                    result.children.push(_this.parseTableCell(c));
                    break;
                case "trPr":
                    _this.parseTableRowProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableRowProperties = function (elem, row) {
        row.cssStyle = this.parseDefaultProperties(elem, {}, null, function (c) {
            switch (c.localName) {
                case "cnfStyle":
                    row.className = values.classNameOfCnfStyle(c);
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseTableCell = function (node) {
        var _this = this;
        var result = { type: dom_1.DomType.Cell, children: [] };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "tbl":
                    result.children.push(_this.parseTable(c));
                    break;
                case "p":
                    result.children.push(_this.parseParagraph(c));
                    break;
                case "tcPr":
                    _this.parseTableCellProperties(c, result);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseTableCellProperties = function (elem, cell) {
        cell.cssStyle = this.parseDefaultProperties(elem, {}, null, function (c) {
            switch (c.localName) {
                case "gridSpan":
                    cell.span = xml.intAttr(c, "val", null);
                    break;
                case "vMerge":
                    break;
                case "cnfStyle":
                    cell.className = values.classNameOfCnfStyle(c);
                    break;
                default:
                    return false;
            }
            return true;
        });
    };
    DocumentParser.prototype.parseDefaultProperties = function (elem, style, childStyle, handler) {
        var _this = this;
        if (style === void 0) { style = null; }
        if (childStyle === void 0) { childStyle = null; }
        if (handler === void 0) { handler = null; }
        style = style || {};
        xml.foreach(elem, function (c) {
            switch (c.localName) {
                case "jc":
                    style["text-align"] = values.valueOfJc(c);
                    break;
                case "textAlignment":
                    style["vertical-align"] = values.valueOfTextAlignment(c);
                    break;
                case "color":
                    style["color"] = xml.colorAttr(c, "val", null, exports.autos.color);
                    break;
                case "sz":
                    style["font-size"] = style["min-height"] = xml.sizeAttr(c, "val", SizeType.FontSize);
                    break;
                case "shd":
                    style["background-color"] = xml.colorAttr(c, "fill", null, exports.autos.shd);
                    break;
                case "highlight":
                    style["background-color"] = xml.colorAttr(c, "val", null, exports.autos.highlight);
                    break;
                case "tcW":
                    if (_this.options.ignoreWidth)
                        break;
                case "tblW":
                    style["width"] = values.valueOfSize(c, "w");
                    break;
                case "trHeight":
                    _this.parseTrHeight(c, style);
                    break;
                case "strike":
                    style["text-decoration"] = values.valueOfStrike(c);
                    break;
                case "b":
                    style["font-weight"] = xml.boolAttr(c, "val", true) ? "bold" : "normal";
                    break;
                case "i":
                    style["font-style"] = xml.boolAttr(c, "val", true) ? "italic" : "normal";
                    break;
                case "caps":
                    style["text-transform"] = xml.boolAttr(c, "val", true) ? "uppercase" : "none";
                    break;
                case "smallCaps":
                    style["text-transform"] = xml.boolAttr(c, "val", true) ? "lowercase" : "none";
                    break;
                case "u":
                    _this.parseUnderline(c, style);
                    break;
                case "ind":
                case "tblInd":
                    _this.parseIndentation(c, style);
                    break;
                case "rFonts":
                    _this.parseFont(c, style);
                    break;
                case "tblBorders":
                    _this.parseBorderProperties(c, childStyle || style);
                    break;
                case "tblCellSpacing":
                    style["border-spacing"] = values.valueOfMargin(c);
                    style["border-collapse"] = "separate";
                    break;
                case "pBdr":
                    _this.parseBorderProperties(c, style);
                    break;
                case "bdr":
                    style["border"] = values.valueOfBorder(c);
                    break;
                case "tcBorders":
                    _this.parseBorderProperties(c, style);
                    break;
                case "noWrap":
                    break;
                case "tblCellMar":
                case "tcMar":
                    _this.parseMarginProperties(c, childStyle || style);
                    break;
                case "tblLayout":
                    style["table-layout"] = values.valueOfTblLayout(c);
                    break;
                case "vAlign":
                    style["vertical-align"] = values.valueOfTextAlignment(c);
                    break;
                case "spacing":
                    if (elem.localName == "pPr")
                        _this.parseSpacing(c, style);
                    break;
                case "lang":
                case "noProof":
                case "webHidden":
                    break;
                default:
                    if (handler != null && !handler(c))
                        _this.options.debug && console.warn("DOCX: Unknown document element: ".concat(c.localName));
                    break;
            }
        });
        return style;
    };
    DocumentParser.prototype.parseUnderline = function (node, style) {
        var val = xml.stringAttr(node, "val");
        if (val == null || val == "none")
            return;
        switch (val) {
            case "dash":
            case "dashDotDotHeavy":
            case "dashDotHeavy":
            case "dashedHeavy":
            case "dashLong":
            case "dashLongHeavy":
            case "dotDash":
            case "dotDotDash":
                style["text-decoration-style"] = "dashed";
                break;
            case "dotted":
            case "dottedHeavy":
                style["text-decoration-style"] = "dotted";
                break;
            case "double":
                style["text-decoration-style"] = "double";
                break;
            case "single":
            case "thick":
                style["text-decoration"] = "underline";
                break;
            case "wave":
            case "wavyDouble":
            case "wavyHeavy":
                style["text-decoration-style"] = "wavy";
                break;
            case "words":
                style["text-decoration"] = "underline";
                break;
        }
        var col = xml.colorAttr(node, "color");
        if (col)
            style["text-decoration-color"] = col;
    };
    DocumentParser.prototype.parseFont = function (node, style) {
        var ascii = xml.stringAttr(node, "ascii");
        if (ascii)
            style["font-family"] = ascii;
    };
    DocumentParser.prototype.parseIndentation = function (node, style) {
        var firstLine = xml.sizeAttr(node, "firstLine");
        var hanging = xml.sizeAttr(node, "hanging");
        var left = xml.sizeAttr(node, "left");
        var start = xml.sizeAttr(node, "start");
        var right = xml.sizeAttr(node, "right");
        var end = xml.sizeAttr(node, "end");
        if (firstLine)
            style["text-indent"] = firstLine;
        if (hanging)
            style["text-indent"] = "-".concat(hanging);
        if (left || start)
            style["margin-left"] = left || start;
        if (right || end)
            style["margin-right"] = right || end;
    };
    DocumentParser.prototype.parseSpacing = function (node, style) {
        var before = xml.sizeAttr(node, "before");
        var after = xml.sizeAttr(node, "after");
        var line = xml.intAttr(node, "line", null);
        var lineRule = xml.stringAttr(node, "lineRule");
        if (before)
            style["margin-top"] = before;
        if (after)
            style["margin-bottom"] = after;
        if (line !== null) {
            switch (lineRule) {
                case "auto":
                    style["line-height"] = "".concat((line / 240).toFixed(2));
                    break;
                case "atLeast":
                    style["line-height"] = "calc(100% + ".concat(line / 20, "pt)");
                    break;
                default:
                    style["line-height"] = style["min-height"] = "".concat(line / 20, "pt");
                    break;
            }
        }
    };
    DocumentParser.prototype.parseMarginProperties = function (node, output) {
        xml.foreach(node, function (c) {
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
    };
    DocumentParser.prototype.parseTrHeight = function (node, output) {
        switch (xml.stringAttr(node, "hRule")) {
            case "exact":
                output["height"] = xml.sizeAttr(node, "val");
                break;
            case "atLeast":
            default:
                output["height"] = xml.sizeAttr(node, "val");
                break;
        }
    };
    DocumentParser.prototype.parseBorderProperties = function (node, output) {
        xml.foreach(node, function (c) {
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
    };
    return DocumentParser;
}());
exports.DocumentParser = DocumentParser;
var SizeType;
(function (SizeType) {
    SizeType[SizeType["FontSize"] = 0] = "FontSize";
    SizeType[SizeType["Dxa"] = 1] = "Dxa";
    SizeType[SizeType["Emu"] = 2] = "Emu";
    SizeType[SizeType["Border"] = 3] = "Border";
    SizeType[SizeType["Percent"] = 4] = "Percent";
})(SizeType || (SizeType = {}));
var xml = (function () {
    function xml() {
    }
    xml.foreach = function (node, cb) {
        for (var i = 0; i < node.childNodes.length; i++) {
            var n = node.childNodes[i];
            if (n.nodeType == 1)
                cb(n);
        }
    };
    xml.elementStringAttr = function (elem, nodeName, attrName) {
        var n = xml_parser_1.default.element(elem, nodeName);
        return n ? xml.stringAttr(n, attrName) : null;
    };
    xml.stringAttr = function (node, attrName) {
        return xml_parser_1.default.attr(node, attrName);
    };
    xml.colorAttr = function (node, attrName, defValue, autoColor) {
        if (defValue === void 0) { defValue = null; }
        if (autoColor === void 0) { autoColor = 'black'; }
        var v = xml.stringAttr(node, attrName);
        switch (v) {
            case "yellow":
                return v;
            case "auto":
                return autoColor;
        }
        return v ? "#".concat(v) : defValue;
    };
    xml.boolAttr = function (node, attrName, defValue) {
        if (defValue === void 0) { defValue = false; }
        return xml_parser_1.default.boolAttr(node, attrName, defValue);
    };
    xml.intAttr = function (node, attrName, defValue) {
        if (defValue === void 0) { defValue = 0; }
        var val = xml.stringAttr(node, attrName);
        return val ? parseInt(xml.stringAttr(node, attrName)) : defValue;
    };
    xml.sizeAttr = function (node, attrName, type) {
        if (type === void 0) { type = SizeType.Dxa; }
        return xml.convertSize(xml.stringAttr(node, attrName), type);
    };
    xml.sizeValue = function (node, type) {
        if (type === void 0) { type = SizeType.Dxa; }
        return xml.convertSize(node.textContent, type);
    };
    xml.convertSize = function (val, type) {
        if (type === void 0) { type = SizeType.Dxa; }
        if (val == null || /.+p[xt]$/.test(val))
            return val;
        var intVal = parseInt(val);
        switch (type) {
            case SizeType.Dxa: return (0.05 * intVal).toFixed(2) + "pt";
            case SizeType.Emu: return (intVal / 12700).toFixed(2) + "pt";
            case SizeType.FontSize: return (0.5 * intVal).toFixed(2) + "pt";
            case SizeType.Border: return (0.125 * intVal).toFixed(2) + "pt";
            case SizeType.Percent: return (0.02 * intVal).toFixed(2) + "%";
        }
        return val;
    };
    xml.className = function (node, attrName) {
        var val = xml.stringAttr(node, attrName);
        return val && val.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and');
    };
    return xml;
}());
var values = (function () {
    function values() {
    }
    values.valueOfBold = function (c) {
        return xml.boolAttr(c, "val", true) ? "bold" : "normal";
    };
    values.valueOfSize = function (c, attr) {
        var type = SizeType.Dxa;
        switch (xml.stringAttr(c, "type")) {
            case "dxa": break;
            case "pct":
                type = SizeType.Percent;
                break;
            case "auto": return "auto";
        }
        return xml.sizeAttr(c, attr, type);
    };
    values.valueOfStrike = function (c) {
        return xml.boolAttr(c, "val", true) ? "line-through" : "none";
    };
    values.valueOfMargin = function (c) {
        return xml.sizeAttr(c, "w");
    };
    values.valueOfBorder = function (c) {
        var type = xml.stringAttr(c, "val");
        if (type == "nil")
            return "none";
        var color = xml.colorAttr(c, "color");
        var size = xml.sizeAttr(c, "sz", SizeType.Border);
        return "".concat(size, " solid ").concat(color == "auto" ? "black" : color);
    };
    values.valueOfTblLayout = function (c) {
        var type = xml.stringAttr(c, "val");
        return type == "fixed" ? "fixed" : "auto";
    };
    values.classNameOfCnfStyle = function (c) {
        var className = "";
        var val = xml.stringAttr(c, "val");
        if (val[0] == "1")
            className += " first-row";
        if (val[1] == "1")
            className += " last-row";
        if (val[2] == "1")
            className += " first-col";
        if (val[3] == "1")
            className += " last-col";
        if (val[4] == "1")
            className += " odd-col";
        if (val[5] == "1")
            className += " even-col";
        if (val[6] == "1")
            className += " odd-row";
        if (val[7] == "1")
            className += " even-row";
        if (val[8] == "1")
            className += " ne-cell";
        if (val[9] == "1")
            className += " nw-cell";
        if (val[10] == "1")
            className += " se-cell";
        if (val[11] == "1")
            className += " sw-cell";
        return className.trim();
    };
    values.valueOfJc = function (c) {
        var type = xml.stringAttr(c, "val");
        switch (type) {
            case "start":
            case "left": return "left";
            case "center": return "center";
            case "end":
            case "right": return "right";
            case "both": return "justify";
        }
        return type;
    };
    values.valueOfTextAlignment = function (c) {
        var type = xml.stringAttr(c, "val");
        switch (type) {
            case "auto":
            case "baseline": return "baseline";
            case "top": return "top";
            case "center": return "middle";
            case "bottom": return "bottom";
        }
        return type;
    };
    values.addSize = function (a, b) {
        if (a == null)
            return b;
        if (b == null)
            return a;
        return "calc(".concat(a, " + ").concat(b, ")");
    };
    values.checkMask = function (num, mask) {
        return (num & mask) == mask;
    };
    values.classNameOftblLook = function (c) {
        var className = "";
        if (xml.boolAttr(c, "firstColumn"))
            className += " first-col";
        if (xml.boolAttr(c, "firstRow"))
            className += " first-row";
        if (xml.boolAttr(c, "lastColumn"))
            className += " lat-col";
        if (xml.boolAttr(c, "lastRow"))
            className += " last-row";
        if (xml.boolAttr(c, "noHBand"))
            className += " no-hband";
        if (xml.boolAttr(c, "noVBand"))
            className += " no-vband";
        return className.trim();
    };
    return values;
}());


/***/ }),

/***/ "./src/document-props/core-props-part.ts":
/*!***********************************************!*\
  !*** ./src/document-props/core-props-part.ts ***!
  \***********************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.CorePropsPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var core_props_1 = __webpack_require__(/*! ./core-props */ "./src/document-props/core-props.ts");
var CorePropsPart = (function (_super) {
    __extends(CorePropsPart, _super);
    function CorePropsPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    CorePropsPart.prototype.parseXml = function (root) {
        this.props = (0, core_props_1.parseCoreProps)(root, this._package.xmlParser);
    };
    return CorePropsPart;
}(part_1.Part));
exports.CorePropsPart = CorePropsPart;


/***/ }),

/***/ "./src/document-props/core-props.ts":
/*!******************************************!*\
  !*** ./src/document-props/core-props.ts ***!
  \******************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseCoreProps = void 0;
function parseCoreProps(root, xmlParser) {
    var result = {};
    for (var _i = 0, _a = xmlParser.elements(root); _i < _a.length; _i++) {
        var el = _a[_i];
        switch (el.localName) {
            case "title":
                result.title = el.textContent;
                break;
            case "description":
                result.description = el.textContent;
                break;
            case "subject":
                result.subject = el.textContent;
                break;
            case "creator":
                result.creator = el.textContent;
                break;
            case "keywords":
                result.keywords = el.textContent;
                break;
            case "language":
                result.language = el.textContent;
                break;
            case "lastModifiedBy":
                result.lastModifiedBy = el.textContent;
                break;
            case "revision":
                el.textContent && (result.revision = parseInt(el.textContent));
                break;
        }
    }
    return result;
}
exports.parseCoreProps = parseCoreProps;


/***/ }),

/***/ "./src/document-props/extended-props-part.ts":
/*!***************************************************!*\
  !*** ./src/document-props/extended-props-part.ts ***!
  \***************************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.ExtendedPropsPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var extended_props_1 = __webpack_require__(/*! ./extended-props */ "./src/document-props/extended-props.ts");
var ExtendedPropsPart = (function (_super) {
    __extends(ExtendedPropsPart, _super);
    function ExtendedPropsPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ExtendedPropsPart.prototype.parseXml = function (root) {
        this.props = (0, extended_props_1.parseExtendedProps)(root, this._package.xmlParser);
    };
    return ExtendedPropsPart;
}(part_1.Part));
exports.ExtendedPropsPart = ExtendedPropsPart;


/***/ }),

/***/ "./src/document-props/extended-props.ts":
/*!**********************************************!*\
  !*** ./src/document-props/extended-props.ts ***!
  \**********************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseExtendedProps = void 0;
function parseExtendedProps(root, xmlParser) {
    var result = {};
    for (var _i = 0, _a = xmlParser.elements(root); _i < _a.length; _i++) {
        var el = _a[_i];
        switch (el.localName) {
            case "Template":
                result.template = el.textContent;
                break;
            case "Pages":
                result.pages = safeParseToInt(el.textContent);
                break;
            case "Words":
                result.words = safeParseToInt(el.textContent);
                break;
            case "Characters":
                result.characters = safeParseToInt(el.textContent);
                break;
            case "Application":
                result.application = el.textContent;
                break;
            case "Lines":
                result.lines = safeParseToInt(el.textContent);
                break;
            case "Paragraphs":
                result.paragraphs = safeParseToInt(el.textContent);
                break;
            case "Company":
                result.company = el.textContent;
                break;
            case "AppVersion":
                result.appVersion = el.textContent;
                break;
        }
    }
    return result;
}
exports.parseExtendedProps = parseExtendedProps;
function safeParseToInt(value) {
    if (typeof value === 'undefined')
        return;
    return parseInt(value);
}


/***/ }),

/***/ "./src/document/bookmark.ts":
/*!**********************************!*\
  !*** ./src/document/bookmark.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseBookmarkEnd = exports.parseBookmarkStart = void 0;
var dom_1 = __webpack_require__(/*! ./dom */ "./src/document/dom.ts");
function parseBookmarkStart(elem, xml) {
    return {
        type: dom_1.DomType.BookmarkStart,
        id: xml.attr(elem, "id"),
        name: xml.attr(elem, "name"),
        colFirst: xml.intAttr(elem, "colFirst"),
        colLast: xml.intAttr(elem, "colLast")
    };
}
exports.parseBookmarkStart = parseBookmarkStart;
function parseBookmarkEnd(elem, xml) {
    return {
        type: dom_1.DomType.BookmarkEnd,
        id: xml.attr(elem, "id")
    };
}
exports.parseBookmarkEnd = parseBookmarkEnd;


/***/ }),

/***/ "./src/document/common.ts":
/*!********************************!*\
  !*** ./src/document/common.ts ***!
  \********************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseCommonProperty = exports.convertLength = exports.LengthUsage = exports.ns = void 0;
exports.ns = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
};
exports.LengthUsage = {
    Dxa: { mul: 0.05, unit: "pt" },
    Emu: { mul: 1 / 12700, unit: "pt" },
    FontSize: { mul: 0.5, unit: "pt" },
    Border: { mul: 0.125, unit: "pt" },
    Point: { mul: 1, unit: "pt" },
    Percent: { mul: 0.02, unit: "%" },
    LineHeight: { mul: 1 / 240, unit: null }
};
function convertLength(val, usage) {
    if (usage === void 0) { usage = exports.LengthUsage.Dxa; }
    return val ? { value: parseInt(val) * usage.mul, type: usage.unit } : null;
}
exports.convertLength = convertLength;
function parseCommonProperty(elem, props, xml) {
    if (elem.namespaceURI != exports.ns.wordml)
        return false;
    switch (elem.localName) {
        case "color":
            props.color = xml.attr(elem, "val");
            break;
        case "sz":
            props.fontSize = xml.lengthAttr(elem, "val", exports.LengthUsage.FontSize);
            break;
        default:
            return false;
    }
    return true;
}
exports.parseCommonProperty = parseCommonProperty;


/***/ }),

/***/ "./src/document/document-part.ts":
/*!***************************************!*\
  !*** ./src/document/document-part.ts ***!
  \***************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.DocumentPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var DocumentPart = (function (_super) {
    __extends(DocumentPart, _super);
    function DocumentPart(pkg, path, parser) {
        var _this = _super.call(this, pkg, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    DocumentPart.prototype.parseXml = function (root) {
        this.body = this._documentParser.parseDocumentFile(root);
    };
    return DocumentPart;
}(part_1.Part));
exports.DocumentPart = DocumentPart;


/***/ }),

/***/ "./src/document/dom.ts":
/*!*****************************!*\
  !*** ./src/document/dom.ts ***!
  \*****************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.DomType = void 0;
var DomType;
(function (DomType) {
    DomType["Document"] = "document";
    DomType["Paragraph"] = "paragraph";
    DomType["Run"] = "run";
    DomType["Break"] = "break";
    DomType["Table"] = "table";
    DomType["Row"] = "row";
    DomType["Cell"] = "cell";
    DomType["Hyperlink"] = "hyperlink";
    DomType["Drawing"] = "drawing";
    DomType["Image"] = "image";
    DomType["Text"] = "text";
    DomType["Tab"] = "tab";
    DomType["Symbol"] = "symbol";
    DomType["BookmarkStart"] = "bookmarkStart";
    DomType["BookmarkEnd"] = "bookmarkEnd";
    DomType["Footer"] = "footer";
    DomType["Header"] = "header";
    DomType["FootnoteReference"] = "footnoteReference";
    DomType["Footnote"] = "footnote";
})(DomType = exports.DomType || (exports.DomType = {}));


/***/ }),

/***/ "./src/document/line-spacing.ts":
/*!**************************************!*\
  !*** ./src/document/line-spacing.ts ***!
  \**************************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseLineSpacing = void 0;
function parseLineSpacing(elem, xml) {
    return {
        before: xml.lengthAttr(elem, "before"),
        after: xml.lengthAttr(elem, "after"),
        line: xml.intAttr(elem, "line"),
        lineRule: xml.attr(elem, "lineRule")
    };
}
exports.parseLineSpacing = parseLineSpacing;


/***/ }),

/***/ "./src/document/paragraph.ts":
/*!***********************************!*\
  !*** ./src/document/paragraph.ts ***!
  \***********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseNumbering = exports.parseTabs = exports.parseParagraphProperty = exports.parseParagraphProperties = void 0;
var common_1 = __webpack_require__(/*! ./common */ "./src/document/common.ts");
var section_1 = __webpack_require__(/*! ./section */ "./src/document/section.ts");
var line_spacing_1 = __webpack_require__(/*! ./line-spacing */ "./src/document/line-spacing.ts");
var run_1 = __webpack_require__(/*! ./run */ "./src/document/run.ts");
function parseParagraphProperties(elem, xml) {
    var result = {};
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var el = _a[_i];
        parseParagraphProperty(el, result, xml);
    }
    return result;
}
exports.parseParagraphProperties = parseParagraphProperties;
function parseParagraphProperty(elem, props, xml) {
    if (elem.namespaceURI != common_1.ns.wordml)
        return false;
    if ((0, common_1.parseCommonProperty)(elem, props, xml))
        return true;
    switch (elem.localName) {
        case "tabs":
            props.tabs = parseTabs(elem, xml);
            break;
        case "sectPr":
            props.sectionProps = (0, section_1.parseSectionProperties)(elem, xml);
            break;
        case "numPr":
            props.numbering = parseNumbering(elem, xml);
            break;
        case "spacing":
            props.lineSpacing = (0, line_spacing_1.parseLineSpacing)(elem, xml);
            return false;
            break;
        case "textAlignment":
            props.textAlignment = xml.attr(elem, "val");
            return false;
            break;
        case "keepNext":
            props.keepLines = xml.boolAttr(elem, "val", true);
            break;
        case "keepNext":
            props.keepNext = xml.boolAttr(elem, "val", true);
            break;
        case "pageBreakBefore":
            props.pageBreakBefore = xml.boolAttr(elem, "val", true);
            break;
        case "outlineLvl":
            props.outlineLevel = xml.intAttr(elem, "val");
            break;
        case "pStyle":
            props.styleName = xml.attr(elem, "val");
            break;
        case "rPr":
            props.runProps = (0, run_1.parseRunProperties)(elem, xml);
            break;
        default:
            return false;
    }
    return true;
}
exports.parseParagraphProperty = parseParagraphProperty;
function parseTabs(elem, xml) {
    return xml.elements(elem, "tab")
        .map(function (e) { return ({
        position: xml.lengthAttr(e, "pos"),
        leader: xml.attr(e, "leader"),
        style: xml.attr(e, "val")
    }); });
}
exports.parseTabs = parseTabs;
function parseNumbering(elem, xml) {
    var result = {};
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "numId":
                result.id = xml.attr(e, "val");
                break;
            case "ilvl":
                result.level = xml.intAttr(e, "val");
                break;
        }
    }
    return result;
}
exports.parseNumbering = parseNumbering;


/***/ }),

/***/ "./src/document/run.ts":
/*!*****************************!*\
  !*** ./src/document/run.ts ***!
  \*****************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseRunProperty = exports.parseRunProperties = void 0;
var common_1 = __webpack_require__(/*! ./common */ "./src/document/common.ts");
function parseRunProperties(elem, xml) {
    var result = {};
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var el = _a[_i];
        parseRunProperty(el, result, xml);
    }
    return result;
}
exports.parseRunProperties = parseRunProperties;
function parseRunProperty(elem, props, xml) {
    if ((0, common_1.parseCommonProperty)(elem, props, xml))
        return true;
    return false;
}
exports.parseRunProperty = parseRunProperty;


/***/ }),

/***/ "./src/document/section.ts":
/*!*********************************!*\
  !*** ./src/document/section.ts ***!
  \*********************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseSectionProperties = exports.SectionType = void 0;
var SectionType;
(function (SectionType) {
    SectionType["Continuous"] = "continuous";
    SectionType["NextPage"] = "nextPage";
    SectionType["NextColumn"] = "nextColumn";
    SectionType["EvenPage"] = "evenPage";
    SectionType["OddPage"] = "oddPage";
})(SectionType = exports.SectionType || (exports.SectionType = {}));
function parseSectionProperties(elem, xml) {
    var _a, _b;
    var section = {};
    for (var _i = 0, _c = xml.elements(elem); _i < _c.length; _i++) {
        var e = _c[_i];
        switch (e.localName) {
            case "pgSz":
                section.pageSize = {
                    width: xml.lengthAttr(e, "w"),
                    height: xml.lengthAttr(e, "h"),
                    orientation: xml.attr(e, "orient")
                };
                break;
            case "type":
                section.type = xml.attr(e, "val");
                break;
            case "pgMar":
                section.pageMargins = {
                    left: xml.lengthAttr(e, "left"),
                    right: xml.lengthAttr(e, "right"),
                    top: xml.lengthAttr(e, "top"),
                    bottom: xml.lengthAttr(e, "bottom"),
                    header: xml.lengthAttr(e, "header"),
                    footer: xml.lengthAttr(e, "footer"),
                    gutter: xml.lengthAttr(e, "gutter"),
                };
                break;
            case "cols":
                section.columns = parseColumns(e, xml);
                break;
            case "headerReference":
                ((_a = section.headerRefs) !== null && _a !== void 0 ? _a : (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml));
                break;
            case "footerReference":
                ((_b = section.footerRefs) !== null && _b !== void 0 ? _b : (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml));
                break;
        }
    }
    return section;
}
exports.parseSectionProperties = parseSectionProperties;
function parseColumns(elem, xml) {
    return {
        numberOfColumns: xml.intAttr(elem, "num"),
        space: xml.lengthAttr(elem, "space"),
        separator: xml.boolAttr(elem, "sep"),
        equalWidth: xml.boolAttr(elem, "equalWidth", true),
        columns: xml.elements(elem, "col")
            .map(function (e) { return ({
            width: xml.lengthAttr(e, "w"),
            space: xml.lengthAttr(e, "space")
        }); })
    };
}
function parseFooterHeaderReference(elem, xml) {
    return {
        id: xml.attr(elem, "id"),
        type: xml.attr(elem, "type"),
    };
}


/***/ }),

/***/ "./src/docx-preview.ts":
/*!*****************************!*\
  !*** ./src/docx-preview.ts ***!
  \*****************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.renderAsync = exports.praseAsync = exports.defaultOptions = void 0;
var word_document_1 = __webpack_require__(/*! ./word-document */ "./src/word-document.ts");
var document_parser_1 = __webpack_require__(/*! ./document-parser */ "./src/document-parser.ts");
var html_renderer_1 = __webpack_require__(/*! ./html-renderer */ "./src/html-renderer.ts");
exports.defaultOptions = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true
};
function praseAsync(data, userOptions) {
    if (userOptions === void 0) { userOptions = null; }
    var ops = __assign(__assign({}, exports.defaultOptions), userOptions);
    return word_document_1.WordDocument.load(data, new document_parser_1.DocumentParser(ops), ops);
}
exports.praseAsync = praseAsync;
function renderAsync(data, bodyContainer, styleContainer, userOptions) {
    if (styleContainer === void 0) { styleContainer = null; }
    if (userOptions === void 0) { userOptions = null; }
    var ops = __assign(__assign({}, exports.defaultOptions), userOptions);
    var renderer = new html_renderer_1.HtmlRenderer(window.document);
    return word_document_1.WordDocument
        .load(data, new document_parser_1.DocumentParser(ops), ops)
        .then(function (doc) {
        renderer.render(doc, bodyContainer, styleContainer, ops);
        return doc;
    });
}
exports.renderAsync = renderAsync;


/***/ }),

/***/ "./src/font-table/font-table.ts":
/*!**************************************!*\
  !*** ./src/font-table/font-table.ts ***!
  \**************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.FontTablePart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var fonts_1 = __webpack_require__(/*! ./fonts */ "./src/font-table/fonts.ts");
var FontTablePart = (function (_super) {
    __extends(FontTablePart, _super);
    function FontTablePart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    FontTablePart.prototype.parseXml = function (root) {
        this.fonts = (0, fonts_1.parseFonts)(root, this._package.xmlParser);
    };
    return FontTablePart;
}(part_1.Part));
exports.FontTablePart = FontTablePart;


/***/ }),

/***/ "./src/font-table/fonts.ts":
/*!*********************************!*\
  !*** ./src/font-table/fonts.ts ***!
  \*********************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseFont = exports.parseFonts = void 0;
function parseFonts(root, xmlParser) {
    return xmlParser.elements(root).map(function (el) { return parseFont(el, xmlParser); });
}
exports.parseFonts = parseFonts;
function parseFont(elem, xmlParser) {
    var result = {
        name: xmlParser.attr(elem, "name")
    };
    for (var _i = 0, _a = xmlParser.elements(elem); _i < _a.length; _i++) {
        var el = _a[_i];
        switch (el.localName) {
            case "family":
                result.family = xmlParser.attr(el, "val");
                break;
            case "altName":
                result.altName = xmlParser.attr(el, "val");
                break;
            case "embedRegular":
                result.fontKey = xmlParser.attr(el, "fontKey");
                result.refId = xmlParser.attr(el, "id");
                break;
        }
    }
    return result;
}
exports.parseFont = parseFont;


/***/ }),

/***/ "./src/footer/footer-part.ts":
/*!***********************************!*\
  !*** ./src/footer/footer-part.ts ***!
  \***********************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.FooterPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var FooterPart = (function (_super) {
    __extends(FooterPart, _super);
    function FooterPart(pkg, path, parser) {
        var _this = _super.call(this, pkg, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    FooterPart.prototype.parseXml = function (root) {
        this.footerElement = this._documentParser.parseFooter(root);
    };
    return FooterPart;
}(part_1.Part));
exports.FooterPart = FooterPart;


/***/ }),

/***/ "./src/footer/footer.ts":
/*!******************************!*\
  !*** ./src/footer/footer.ts ***!
  \******************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.WmlFooter = void 0;
var dom_1 = __webpack_require__(/*! ../document/dom */ "./src/document/dom.ts");
var WmlFooter = (function () {
    function WmlFooter() {
        this.type = dom_1.DomType.Footer;
        this.children = [];
        this.cssStyle = {};
    }
    return WmlFooter;
}());
exports.WmlFooter = WmlFooter;


/***/ }),

/***/ "./src/footnotes/footnote.ts":
/*!***********************************!*\
  !*** ./src/footnotes/footnote.ts ***!
  \***********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.WmlFootnote = void 0;
var dom_1 = __webpack_require__(/*! ../document/dom */ "./src/document/dom.ts");
var WmlFootnote = (function () {
    function WmlFootnote() {
        this.type = dom_1.DomType.Footnote;
        this.children = [];
        this.cssStyle = {};
    }
    return WmlFootnote;
}());
exports.WmlFootnote = WmlFootnote;


/***/ }),

/***/ "./src/footnotes/footnotes-part.ts":
/*!*****************************************!*\
  !*** ./src/footnotes/footnotes-part.ts ***!
  \*****************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.FootnotesPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var FootnotesPart = (function (_super) {
    __extends(FootnotesPart, _super);
    function FootnotesPart(pkg, path, parser) {
        var _this = _super.call(this, pkg, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    FootnotesPart.prototype.parseXml = function (root) {
        this.footnotes = this._documentParser.parseFootnotes(root);
    };
    return FootnotesPart;
}(part_1.Part));
exports.FootnotesPart = FootnotesPart;


/***/ }),

/***/ "./src/header/header-part.ts":
/*!***********************************!*\
  !*** ./src/header/header-part.ts ***!
  \***********************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.HeaderPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var HeaderPart = (function (_super) {
    __extends(HeaderPart, _super);
    function HeaderPart(pkg, path, parser) {
        var _this = _super.call(this, pkg, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    HeaderPart.prototype.parseXml = function (root) {
        this.headerElement = this._documentParser.parseHeader(root);
    };
    return HeaderPart;
}(part_1.Part));
exports.HeaderPart = HeaderPart;


/***/ }),

/***/ "./src/header/header.ts":
/*!******************************!*\
  !*** ./src/header/header.ts ***!
  \******************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.WmlHeader = void 0;
var dom_1 = __webpack_require__(/*! ../document/dom */ "./src/document/dom.ts");
var WmlHeader = (function () {
    function WmlHeader() {
        this.type = dom_1.DomType.Header;
        this.children = [];
        this.cssStyle = {};
    }
    return WmlHeader;
}());
exports.WmlHeader = WmlHeader;


/***/ }),

/***/ "./src/html-renderer.ts":
/*!******************************!*\
  !*** ./src/html-renderer.ts ***!
  \******************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.HtmlRenderer = void 0;
var dom_1 = __webpack_require__(/*! ./document/dom */ "./src/document/dom.ts");
var utils_1 = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var javascript_1 = __webpack_require__(/*! ./javascript */ "./src/javascript.ts");
var HtmlRenderer = (function () {
    function HtmlRenderer(htmlDocument) {
        this.htmlDocument = htmlDocument;
        this.className = "docx";
        this.footnoteMap = {};
    }
    HtmlRenderer.prototype.render = function (document, bodyContainer, styleContainer, options) {
        if (styleContainer === void 0) { styleContainer = null; }
        this.document = document;
        this.options = options;
        this.className = options.className;
        this.styleMap = null;
        styleContainer = styleContainer || bodyContainer;
        removeAllElements(styleContainer);
        removeAllElements(bodyContainer);
        appendComment(styleContainer, "docxjs library predefined styles");
        styleContainer.appendChild(this.renderDefaultStyle());
        if (document.stylesPart != null) {
            this.styleMap = this.processStyles(document.stylesPart.styles);
            appendComment(styleContainer, "docx document styles");
            styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
        }
        if (document.numberingPart) {
            appendComment(styleContainer, "docx document numbering styles");
            styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
        }
        if (document.footnotesPart) {
            this.footnoteMap = (0, utils_1.keyBy)(document.footnotesPart.footnotes, function (x) { return x.id; });
        }
        if (!options.ignoreFonts && document.fontTablePart)
            this.renderFontTable(document.fontTablePart, styleContainer);
        var sectionElements = this.renderSections(document.documentPart.body);
        if (this.options.inWrapper) {
            var wrapper = this.renderWrapper();
            appentElements(wrapper, sectionElements);
            bodyContainer.appendChild(wrapper);
        }
        else {
            appentElements(bodyContainer, sectionElements);
        }
    };
    HtmlRenderer.prototype.renderFontTable = function (fontsPart, styleContainer) {
        var _loop_1 = function (f) {
            this_1.document.loadFont(f.refId, f.fontKey).then(function (fontData) {
                var cssTest = "@font-face {\n                    font-family: \"".concat(f.name, "\";\n                    src: url(").concat(fontData, ");\n                }");
                appendComment(styleContainer, "Font ".concat(f.name));
                styleContainer.appendChild(createStyleElement(cssTest));
            });
        };
        var this_1 = this;
        for (var _i = 0, _a = fontsPart.fonts.filter(function (x) { return x.refId; }); _i < _a.length; _i++) {
            var f = _a[_i];
            _loop_1(f);
        }
    };
    HtmlRenderer.prototype.processClassName = function (className) {
        if (!className)
            return this.className;
        return "".concat(this.className, "_").concat(className);
    };
    HtmlRenderer.prototype.processStyles = function (styles) {
        var stylesMap = (0, utils_1.keyBy)(styles.filter(function (x) { return x.id != null; }), function (x) { return x.id; });
        for (var _i = 0, _a = styles.filter(function (x) { return x.basedOn; }); _i < _a.length; _i++) {
            var style = _a[_i];
            var baseStyle = stylesMap[style.basedOn];
            if (baseStyle) {
                var _loop_2 = function (styleValues) {
                    baseValues = baseStyle.styles.filter(function (x) { return x.target == styleValues.target; });
                    if (baseValues && baseValues.length > 0)
                        this_2.copyStyleProperties(baseValues[0].values, styleValues.values);
                };
                var this_2 = this, baseValues;
                for (var _b = 0, _c = style.styles; _b < _c.length; _b++) {
                    var styleValues = _c[_b];
                    _loop_2(styleValues);
                }
            }
            else if (this.options.debug)
                console.warn("Can't find base style ".concat(style.basedOn));
        }
        for (var _d = 0, styles_1 = styles; _d < styles_1.length; _d++) {
            var style = styles_1[_d];
            style.cssName = this.processClassName(this.escapeClassName(style.id));
        }
        return stylesMap;
    };
    HtmlRenderer.prototype.processElement = function (element) {
        if (element.children) {
            for (var _i = 0, _a = element.children; _i < _a.length; _i++) {
                var e = _a[_i];
                e.className = this.processClassName(e.className);
                e.parent = element;
                if (e.type == dom_1.DomType.Table) {
                    this.processTable(e);
                }
                else {
                    this.processElement(e);
                }
            }
        }
    };
    HtmlRenderer.prototype.processTable = function (table) {
        for (var _i = 0, _a = table.children; _i < _a.length; _i++) {
            var r = _a[_i];
            for (var _b = 0, _c = r.children; _b < _c.length; _b++) {
                var c = _c[_b];
                c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
                    "border-left", "border-right", "border-top", "border-bottom",
                    "padding-left", "padding-right", "padding-top", "padding-bottom"
                ]);
                this.processElement(c);
            }
        }
    };
    HtmlRenderer.prototype.copyStyleProperties = function (input, output, attrs) {
        if (attrs === void 0) { attrs = null; }
        if (!input)
            return output;
        if (output == null)
            output = {};
        if (attrs == null)
            attrs = Object.getOwnPropertyNames(input);
        for (var _i = 0, attrs_1 = attrs; _i < attrs_1.length; _i++) {
            var key = attrs_1[_i];
            if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                output[key] = input[key];
        }
        return output;
    };
    HtmlRenderer.prototype.createSection = function (className, props) {
        var elem = this.htmlDocument.createElement("section");
        elem.className = className;
        if (props) {
            if (props.pageMargins) {
                elem.style.paddingLeft = this.renderLength(props.pageMargins.left);
                elem.style.paddingRight = this.renderLength(props.pageMargins.right);
                elem.style.paddingTop = this.renderLength(props.pageMargins.top);
                elem.style.paddingBottom = this.renderLength(props.pageMargins.bottom);
            }
            if (props.pageSize) {
                if (!this.options.ignoreWidth)
                    elem.style.width = this.renderLength(props.pageSize.width);
                if (!this.options.ignoreHeight)
                    elem.style.minHeight = this.renderLength(props.pageSize.height);
            }
            if (props.columns && props.columns.numberOfColumns) {
                elem.style.columnCount = "".concat(props.columns.numberOfColumns);
                elem.style.columnGap = this.renderLength(props.columns.space);
                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }
        return elem;
    };
    HtmlRenderer.prototype.renderSections = function (document) {
        var result = [];
        this.processElement(document);
        for (var _i = 0, _a = this.splitBySection(document.children); _i < _a.length; _i++) {
            var section = _a[_i];
            this.currentFootnoteIds = [];
            var props = section.sectProps || document.props;
            var sectionElement = this.createSection(this.className, props);
            this.renderStyleValues(document.cssStyle, sectionElement);
            var headerPart = this.options.renderHeaders ? this.findHeaderFooter(props.headerRefs, result.length) : null;
            var footerPart = this.options.renderFooters ? this.findHeaderFooter(props.footerRefs, result.length) : null;
            headerPart && this.renderElements([headerPart.headerElement], sectionElement);
            var contentElement = this.htmlDocument.createElement("article");
            this.renderElements(section.elements, contentElement);
            sectionElement.appendChild(contentElement);
            if (this.options.renderFootnotes) {
                this.renderFootnotes(this.currentFootnoteIds, sectionElement);
            }
            footerPart && this.renderElements([footerPart.footerElement], sectionElement);
            result.push(sectionElement);
        }
        return result;
    };
    HtmlRenderer.prototype.findHeaderFooter = function (refs, page) {
        var _a, _b;
        var ref = refs ? ((_b = (_a = (page == 0 ? refs.find(function (x) { return x.type == "first"; }) : null)) !== null && _a !== void 0 ? _a : (page % 2 == 0 ? refs.find(function (x) { return x.type == "even"; }) : null)) !== null && _b !== void 0 ? _b : refs.find(function (x) { return x.type == "default"; })) : null;
        if (ref == null)
            return null;
        return this.document.findPartByRelId(ref.id, this.document.documentPart);
    };
    HtmlRenderer.prototype.isPageBreakElement = function (elem) {
        if (elem.type != dom_1.DomType.Break)
            return false;
        if (elem.break == "lastRenderedPageBreak")
            return !this.options.ignoreLastRenderedPageBreak;
        return elem.break == "page";
    };
    HtmlRenderer.prototype.splitBySection = function (elements) {
        var _this = this;
        var _a;
        var current = { sectProps: null, elements: [] };
        var result = [current];
        for (var _i = 0, elements_1 = elements; _i < elements_1.length; _i++) {
            var elem = elements_1[_i];
            if (elem.type == dom_1.DomType.Paragraph) {
                var styleName = elem.styleName;
                var s = this.styleMap && styleName ? this.styleMap[styleName] : null;
                if ((_a = s === null || s === void 0 ? void 0 : s.paragraphProps) === null || _a === void 0 ? void 0 : _a.pageBreakBefore) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }
            }
            current.elements.push(elem);
            if (elem.type == dom_1.DomType.Paragraph) {
                var p = elem;
                var sectProps = p.sectionProps;
                var pBreakIndex = -1;
                var rBreakIndex = -1;
                if (this.options.breakPages && p.children) {
                    pBreakIndex = p.children.findIndex(function (r) {
                        var _a, _b;
                        rBreakIndex = (_b = (_a = r.children) === null || _a === void 0 ? void 0 : _a.findIndex(_this.isPageBreakElement.bind(_this))) !== null && _b !== void 0 ? _b : -1;
                        return rBreakIndex != -1;
                    });
                }
                if (sectProps || pBreakIndex != -1) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }
                if (pBreakIndex != -1) {
                    var breakRun = p.children[pBreakIndex];
                    var splitRun = rBreakIndex < breakRun.children.length - 1;
                    if (pBreakIndex < p.children.length - 1 || splitRun) {
                        var children = elem.children;
                        var newParagraph = __assign(__assign({}, elem), { children: children.slice(pBreakIndex) });
                        elem.children = children.slice(0, pBreakIndex);
                        current.elements.push(newParagraph);
                        if (splitRun) {
                            var runChildren = breakRun.children;
                            var newRun = __assign(__assign({}, breakRun), { children: runChildren.slice(0, rBreakIndex) });
                            elem.children.push(newRun);
                            breakRun.children = runChildren.slice(rBreakIndex);
                        }
                    }
                }
            }
        }
        var currentSectProps = null;
        for (var i = result.length - 1; i >= 0; i--) {
            if (result[i].sectProps == null) {
                result[i].sectProps = currentSectProps;
            }
            else {
                currentSectProps = result[i].sectProps;
            }
        }
        return result;
    };
    HtmlRenderer.prototype.renderLength = function (l) {
        return l ? "".concat(l.value).concat(l.type) : null;
    };
    HtmlRenderer.prototype.renderWrapper = function () {
        var wrapper = document.createElement("div");
        wrapper.className = "".concat(this.className, "-wrapper");
        return wrapper;
    };
    HtmlRenderer.prototype.renderDefaultStyle = function () {
        var c = this.className;
        var styleText = "\n.".concat(c, "-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } \n.").concat(c, "-wrapper>section.").concat(c, " { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }\n.").concat(c, " { color: black; }\nsection.").concat(c, " { box-sizing: border-box; display: flex; flex-flow: column nowrap; }\nsection.").concat(c, ">article { margin-bottom: auto; }\n.").concat(c, " table { border-collapse: collapse; }\n.").concat(c, " table td, .").concat(c, " table th { vertical-align: top; }\n.").concat(c, " p { margin: 0pt; min-height: 1em; }\n.").concat(c, " span { white-space: pre-wrap; }\n");
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderNumbering = function (numberings, styleContainer) {
        var _this = this;
        var styleText = "";
        var rootCounters = [];
        var _loop_3 = function () {
            selector = "p.".concat(this_3.numberingClass(num.id, num.level));
            listStyleType = "none";
            if (num.bullet) {
                var valiable_1 = "--".concat(this_3.className, "-").concat(num.bullet.src).toLowerCase();
                styleText += this_3.styleToString("".concat(selector, ":before"), {
                    "content": "' '",
                    "display": "inline-block",
                    "background": "var(".concat(valiable_1, ")")
                }, num.bullet.style);
                this_3.document.loadNumberingImage(num.bullet.src).then(function (data) {
                    var text = ".".concat(_this.className, "-wrapper { ").concat(valiable_1, ": url(").concat(data, ") }");
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else if (num.levelText) {
                var counter = this_3.numberingCounter(num.id, num.level);
                if (num.level > 0) {
                    styleText += this_3.styleToString("p.".concat(this_3.numberingClass(num.id, num.level - 1)), {
                        "counter-reset": counter
                    });
                }
                else {
                    rootCounters.push(counter);
                }
                styleText += this_3.styleToString("".concat(selector, ":before"), __assign({ "content": this_3.levelTextToContent(num.levelText, num.suff, num.id, this_3.numFormatToCssValue(num.format)), "counter-increment": counter }, num.rStyle));
            }
            else {
                listStyleType = this_3.numFormatToCssValue(num.format);
            }
            styleText += this_3.styleToString(selector, __assign({ "display": "list-item", "list-style-position": "inside", "list-style-type": listStyleType }, num.pStyle));
        };
        var this_3 = this, selector, listStyleType;
        for (var _i = 0, numberings_1 = numberings; _i < numberings_1.length; _i++) {
            var num = numberings_1[_i];
            _loop_3();
        }
        if (rootCounters.length > 0) {
            styleText += this.styleToString(".".concat(this.className, "-wrapper"), {
                "counter-reset": rootCounters.join(" ")
            });
        }
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderStyles = function (styles) {
        var styleText = "";
        var stylesMap = this.styleMap;
        var defautStyles = (0, utils_1.keyBy)(styles.filter(function (s) { return s.isDefault; }), function (s) { return s.target; });
        for (var _i = 0, styles_2 = styles; _i < styles_2.length; _i++) {
            var style = styles_2[_i];
            var subStyles = style.styles;
            if (style.linked) {
                var linkedStyle = style.linked && stylesMap[style.linked];
                if (linkedStyle)
                    subStyles = subStyles.concat(linkedStyle.styles);
                else if (this.options.debug)
                    console.warn("Can't find linked style ".concat(style.linked));
            }
            for (var _a = 0, subStyles_1 = subStyles; _a < subStyles_1.length; _a++) {
                var subStyle = subStyles_1[_a];
                var selector = "";
                if (style.target == subStyle.target)
                    selector += "".concat(style.target, ".").concat(style.cssName);
                else if (style.target)
                    selector += "".concat(style.target, ".").concat(style.cssName, " ").concat(subStyle.target);
                else
                    selector += ".".concat(style.cssName, " ").concat(subStyle.target);
                if (defautStyles[style.target] == style)
                    selector = ".".concat(this.className, " ").concat(style.target, ", ") + selector;
                styleText += this.styleToString(selector, subStyle.values);
            }
        }
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderFootnotes = function (footnoteIds, into) {
        var _this = this;
        var footnotes = footnoteIds.map(function (id) { return _this.footnoteMap[id]; });
        if (footnotes.length > 0) {
            var result = this.htmlDocument.createElement("ol");
            this.renderElements(footnotes, result);
            into.appendChild(result);
        }
    };
    HtmlRenderer.prototype.renderElement = function (elem) {
        switch (elem.type) {
            case dom_1.DomType.Paragraph:
                return this.renderParagraph(elem);
            case dom_1.DomType.BookmarkStart:
                return this.renderBookmarkStart(elem);
            case dom_1.DomType.BookmarkEnd:
                return null;
            case dom_1.DomType.Run:
                return this.renderRun(elem);
            case dom_1.DomType.Table:
                return this.renderTable(elem);
            case dom_1.DomType.Row:
                return this.renderTableRow(elem);
            case dom_1.DomType.Cell:
                return this.renderTableCell(elem);
            case dom_1.DomType.Hyperlink:
                return this.renderHyperlink(elem);
            case dom_1.DomType.Drawing:
                return this.renderDrawing(elem);
            case dom_1.DomType.Image:
                return this.renderImage(elem);
            case dom_1.DomType.Text:
                return this.renderText(elem);
            case dom_1.DomType.Tab:
                return this.renderTab(elem);
            case dom_1.DomType.Symbol:
                return this.renderSymbol(elem);
            case dom_1.DomType.Break:
                return this.renderBreak(elem);
            case dom_1.DomType.Footer:
                return this.renderContainer(elem, "footer");
            case dom_1.DomType.Header:
                return this.renderContainer(elem, "header");
            case dom_1.DomType.Footnote:
                return this.renderContainer(elem, "li");
            case dom_1.DomType.FootnoteReference:
                return this.renderFootnoteReference(elem);
        }
        return null;
    };
    HtmlRenderer.prototype.renderChildren = function (elem, into) {
        return this.renderElements(elem.children, into);
    };
    HtmlRenderer.prototype.renderElements = function (elems, into) {
        var _this = this;
        if (elems == null)
            return null;
        var result = elems.map(function (e) { return _this.renderElement(e); }).filter(function (e) { return e != null; });
        if (into)
            for (var _i = 0, result_1 = result; _i < result_1.length; _i++) {
                var c = result_1[_i];
                into.appendChild(c);
            }
        return result;
    };
    HtmlRenderer.prototype.renderContainer = function (elem, tagName) {
        var result = this.htmlDocument.createElement(tagName);
        this.renderChildren(elem, result);
        return result;
    };
    HtmlRenderer.prototype.renderParagraph = function (elem) {
        var result = this.htmlDocument.createElement("p");
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        this.renderCommonProeprties(result.style, elem);
        if (elem.numbering) {
            var numberingClass = this.numberingClass(elem.numbering.id, elem.numbering.level);
            result.className = (0, utils_1.appendClass)(result.className, numberingClass);
        }
        if (elem.styleName) {
            var styleClassName = this.processClassName(this.escapeClassName(elem.styleName));
            result.className = (0, utils_1.appendClass)(result.className, styleClassName);
        }
        return result;
    };
    HtmlRenderer.prototype.renderRunProperties = function (style, props) {
        this.renderCommonProeprties(style, props);
    };
    HtmlRenderer.prototype.renderCommonProeprties = function (style, props) {
        if (props == null)
            return;
        if (props.color) {
            style["color"] = props.color;
        }
        if (props.fontSize) {
            style["font-size"] = this.renderLength(props.fontSize);
        }
    };
    HtmlRenderer.prototype.renderHyperlink = function (elem) {
        var result = this.htmlDocument.createElement("a");
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        if (elem.href)
            result.href = elem.href;
        return result;
    };
    HtmlRenderer.prototype.renderDrawing = function (elem) {
        var result = this.htmlDocument.createElement("div");
        result.style.display = "inline-block";
        result.style.position = "relative";
        result.style.textIndent = "0px";
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        return result;
    };
    HtmlRenderer.prototype.renderImage = function (elem) {
        var result = this.htmlDocument.createElement("img");
        this.renderStyleValues(elem.cssStyle, result);
        if (this.document) {
            this.document.loadDocumentImage(elem.src).then(function (x) {
                result.src = x;
            });
        }
        return result;
    };
    HtmlRenderer.prototype.renderText = function (elem) {
        return this.htmlDocument.createTextNode(elem.text);
    };
    HtmlRenderer.prototype.renderBreak = function (elem) {
        if (elem.break == "textWrapping") {
            return this.htmlDocument.createElement("br");
        }
        return null;
    };
    HtmlRenderer.prototype.renderSymbol = function (elem) {
        var span = this.htmlDocument.createElement("span");
        span.style.fontFamily = elem.font;
        span.innerHTML = "&#x".concat(elem.char, ";");
        return span;
    };
    HtmlRenderer.prototype.renderFootnoteReference = function (elem) {
        var result = this.htmlDocument.createElement("sup");
        this.currentFootnoteIds.push(elem.id);
        result.textContent = "".concat(this.currentFootnoteIds.length);
        return result;
    };
    HtmlRenderer.prototype.renderTab = function (elem) {
        var tabSpan = this.htmlDocument.createElement("span");
        tabSpan.innerHTML = "&emsp;";
        if (this.options.experimental) {
            setTimeout(function () {
                var paragraph = findParent(elem, dom_1.DomType.Paragraph);
                if ((paragraph === null || paragraph === void 0 ? void 0 : paragraph.tabs) == null)
                    return;
                paragraph.tabs.sort(function (a, b) { return a.position.value - b.position.value; });
                tabSpan.style.display = "inline-block";
                (0, javascript_1.updateTabStop)(tabSpan, paragraph.tabs);
            }, 0);
        }
        return tabSpan;
    };
    HtmlRenderer.prototype.renderBookmarkStart = function (elem) {
        var result = this.htmlDocument.createElement("span");
        result.id = elem.name;
        return result;
    };
    HtmlRenderer.prototype.renderRun = function (elem) {
        if (elem.fldCharType || elem.instrText)
            return null;
        var result = this.htmlDocument.createElement("span");
        if (elem.id)
            result.id = elem.id;
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        if (elem.href) {
            var link = this.htmlDocument.createElement("a");
            link.href = elem.href;
            link.appendChild(result);
            return link;
        }
        else if (elem.wrapper) {
            var wrapper = this.htmlDocument.createElement(elem.wrapper);
            wrapper.appendChild(result);
            return wrapper;
        }
        return result;
    };
    HtmlRenderer.prototype.renderTable = function (elem) {
        var result = this.htmlDocument.createElement("table");
        if (elem.columns)
            result.appendChild(this.renderTableColumns(elem.columns));
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        return result;
    };
    HtmlRenderer.prototype.renderTableColumns = function (columns) {
        var result = this.htmlDocument.createElement("colGroup");
        for (var _i = 0, columns_1 = columns; _i < columns_1.length; _i++) {
            var col = columns_1[_i];
            var colElem = this.htmlDocument.createElement("col");
            if (col.width)
                colElem.style.width = col.width;
            result.appendChild(colElem);
        }
        return result;
    };
    HtmlRenderer.prototype.renderTableRow = function (elem) {
        var result = this.htmlDocument.createElement("tr");
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        return result;
    };
    HtmlRenderer.prototype.renderTableCell = function (elem) {
        var result = this.htmlDocument.createElement("td");
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        if (elem.span)
            result.colSpan = elem.span;
        return result;
    };
    HtmlRenderer.prototype.renderStyleValues = function (style, ouput) {
        if (style == null)
            return;
        for (var key in style) {
            if (style.hasOwnProperty(key)) {
                ouput.style[key] = style[key];
            }
        }
    };
    HtmlRenderer.prototype.renderClass = function (input, ouput) {
        if (input.className)
            ouput.className = input.className;
    };
    HtmlRenderer.prototype.numberingClass = function (id, lvl) {
        return "".concat(this.className, "-num-").concat(id, "-").concat(lvl);
    };
    HtmlRenderer.prototype.styleToString = function (selectors, values, cssText) {
        if (cssText === void 0) { cssText = null; }
        var result = selectors + " {\r\n";
        for (var key in values) {
            result += "  ".concat(key, ": ").concat(values[key], ";\r\n");
        }
        if (cssText)
            result += ";" + cssText;
        return result + "}\r\n";
    };
    HtmlRenderer.prototype.numberingCounter = function (id, lvl) {
        return "".concat(this.className, "-num-").concat(id, "-").concat(lvl);
    };
    HtmlRenderer.prototype.levelTextToContent = function (text, suff, id, numformat) {
        var _this = this;
        var _a;
        var suffMap = {
            "tab": "\\9",
            "space": "\\a0",
        };
        var result = text.replace(/%\d*/g, function (s) {
            var lvl = parseInt(s.substring(1), 10) - 1;
            return "\"counter(".concat(_this.numberingCounter(id, lvl), ", ").concat(numformat, ")\"");
        });
        return "\"".concat(result).concat((_a = suffMap[suff]) !== null && _a !== void 0 ? _a : "", "\"");
    };
    HtmlRenderer.prototype.numFormatToCssValue = function (format) {
        var mapping = {
            "none": "none",
            "bullet": "disc",
            "decimal": "decimal",
            "lowerLetter": "lower-alpha",
            "upperLetter": "upper-alpha",
            "lowerRoman": "lower-roman",
            "upperRoman": "upper-roman",
        };
        return mapping[format] || format;
    };
    HtmlRenderer.prototype.escapeClassName = function (className) {
        return className === null || className === void 0 ? void 0 : className.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and');
    };
    return HtmlRenderer;
}());
exports.HtmlRenderer = HtmlRenderer;
function appentElements(container, children) {
    for (var _i = 0, children_1 = children; _i < children_1.length; _i++) {
        var c = children_1[_i];
        container.appendChild(c);
    }
}
function removeAllElements(elem) {
    while (elem.firstChild) {
        elem.removeChild(elem.firstChild);
    }
}
function createStyleElement(cssText) {
    var styleElement = document.createElement("style");
    styleElement.innerHTML = cssText;
    return styleElement;
}
function appendComment(elem, comment) {
    elem.appendChild(document.createComment(comment));
}
function findParent(elem, type) {
    var parent = elem.parent;
    while (parent != null && parent.type != type)
        parent = parent.parent;
    return parent;
}


/***/ }),

/***/ "./src/javascript.ts":
/*!***************************!*\
  !*** ./src/javascript.ts ***!
  \***************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.updateTabStop = void 0;
function updateTabStop(elem, tabs, pixelToPoint) {
    if (pixelToPoint === void 0) { pixelToPoint = 72 / 96; }
    var p = elem.closest("p");
    var tbb = elem.getBoundingClientRect();
    var pbb = p.getBoundingClientRect();
    var left = (tbb.left - pbb.left) * pixelToPoint;
    var tab = tabs.find(function (t) { return t.style != "clear" && t.position.value > left; });
    if (tab == null)
        return;
    elem.style.display = "inline-block";
    elem.style.width = "".concat((tab.position.value - left), "pt");
    switch (tab.leader) {
        case "dot":
        case "middleDot":
            elem.style.borderBottom = "1px black dotted";
            break;
        case "hyphen":
        case "heavy":
        case "underscore":
            elem.style.borderBottom = "1px black solid";
            break;
    }
}
exports.updateTabStop = updateTabStop;


/***/ }),

/***/ "./src/numbering/numbering-part.ts":
/*!*****************************************!*\
  !*** ./src/numbering/numbering-part.ts ***!
  \*****************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.NumberingPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var numbering_1 = __webpack_require__(/*! ./numbering */ "./src/numbering/numbering.ts");
var NumberingPart = (function (_super) {
    __extends(NumberingPart, _super);
    function NumberingPart(pkg, path, parser) {
        var _this = _super.call(this, pkg, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    NumberingPart.prototype.parseXml = function (root) {
        Object.assign(this, (0, numbering_1.parseNumberingPart)(root, this._package.xmlParser));
        this.domNumberings = this._documentParser.parseNumberingFile(root);
    };
    return NumberingPart;
}(part_1.Part));
exports.NumberingPart = NumberingPart;


/***/ }),

/***/ "./src/numbering/numbering.ts":
/*!************************************!*\
  !*** ./src/numbering/numbering.ts ***!
  \************************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseNumberingBulletPicture = exports.parseNumberingLevelOverrride = exports.parseNumberingLevel = exports.parseAbstractNumbering = exports.parseNumbering = exports.parseNumberingPart = void 0;
var paragraph_1 = __webpack_require__(/*! ../document/paragraph */ "./src/document/paragraph.ts");
var run_1 = __webpack_require__(/*! ../document/run */ "./src/document/run.ts");
function parseNumberingPart(elem, xml) {
    var result = {
        numberings: [],
        abstractNumberings: [],
        bulletPictures: []
    };
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "num":
                result.numberings.push(parseNumbering(e, xml));
                break;
            case "abstractNum":
                result.abstractNumberings.push(parseAbstractNumbering(e, xml));
                break;
            case "numPicBullet":
                result.bulletPictures.push(parseNumberingBulletPicture(e, xml));
                break;
        }
    }
    return result;
}
exports.parseNumberingPart = parseNumberingPart;
function parseNumbering(elem, xml) {
    var result = {
        id: xml.attr(elem, 'numId'),
        overrides: []
    };
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "abstractNumId":
                result.abstractId = xml.attr(e, "val");
                break;
            case "lvlOverride":
                result.overrides.push(parseNumberingLevelOverrride(e, xml));
                break;
        }
    }
    return result;
}
exports.parseNumbering = parseNumbering;
function parseAbstractNumbering(elem, xml) {
    var result = {
        id: xml.attr(elem, 'abstractNumId'),
        levels: []
    };
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "name":
                result.name = xml.attr(e, "val");
                break;
            case "multiLevelType":
                result.multiLevelType = xml.attr(e, "val");
                break;
            case "numStyleLink":
                result.numberingStyleLink = xml.attr(e, "val");
                break;
            case "styleLink":
                result.styleLink = xml.attr(e, "val");
                break;
            case "lvl":
                result.levels.push(parseNumberingLevel(e, xml));
                break;
        }
    }
    return result;
}
exports.parseAbstractNumbering = parseAbstractNumbering;
function parseNumberingLevel(elem, xml) {
    var result = {
        level: xml.intAttr(elem, 'ilvl')
    };
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "start":
                result.start = xml.attr(e, "val");
                break;
            case "lvlRestart":
                result.restart = xml.intAttr(e, "val");
                break;
            case "numFmt":
                result.format = xml.attr(e, "val");
                break;
            case "lvlText":
                result.text = xml.attr(e, "val");
                break;
            case "lvlJc":
                result.justification = xml.attr(e, "val");
                break;
            case "lvlPicBulletId":
                result.bulletPictureId = xml.attr(e, "val");
                break;
            case "pPr":
                result.paragraphProps = (0, paragraph_1.parseParagraphProperties)(e, xml);
                break;
            case "rPr":
                result.runProps = (0, run_1.parseRunProperties)(e, xml);
                break;
        }
    }
    return result;
}
exports.parseNumberingLevel = parseNumberingLevel;
function parseNumberingLevelOverrride(elem, xml) {
    var result = {
        level: xml.intAttr(elem, 'ilvl')
    };
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "startOverride":
                result.start = xml.intAttr(e, "val");
                break;
            case "lvl":
                result.numberingLevel = parseNumberingLevel(e, xml);
                break;
        }
    }
    return result;
}
exports.parseNumberingLevelOverrride = parseNumberingLevelOverrride;
function parseNumberingBulletPicture(elem, xml) {
    var pict = xml.element(elem, "pict");
    var shape = pict && xml.element(pict, "shape");
    var imagedata = shape && xml.element(shape, "imagedata");
    return imagedata ? {
        id: xml.attr(elem, "numPicBulletId"),
        referenceId: xml.attr(imagedata, "id"),
        style: xml.attr(shape, "style")
    } : null;
}
exports.parseNumberingBulletPicture = parseNumberingBulletPicture;


/***/ }),

/***/ "./src/parser/xml-parser.ts":
/*!**********************************!*\
  !*** ./src/parser/xml-parser.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.XmlParser = exports.serializeXmlString = exports.parseXmlString = void 0;
var common_1 = __webpack_require__(/*! ../document/common */ "./src/document/common.ts");
function parseXmlString(xmlString, trimXmlDeclaration) {
    if (trimXmlDeclaration === void 0) { trimXmlDeclaration = false; }
    if (trimXmlDeclaration)
        xmlString = xmlString.replace(/<[?].*[?]>/, "");
    var result = new DOMParser().parseFromString(xmlString, "application/xml");
    var errorText = hasXmlParserError(result);
    if (errorText)
        throw new Error(errorText);
    return result;
}
exports.parseXmlString = parseXmlString;
function hasXmlParserError(doc) {
    var _a;
    return (_a = doc.getElementsByTagName("parsererror")[0]) === null || _a === void 0 ? void 0 : _a.textContent;
}
function serializeXmlString(elem) {
    return new XMLSerializer().serializeToString(elem);
}
exports.serializeXmlString = serializeXmlString;
var XmlParser = (function () {
    function XmlParser() {
    }
    XmlParser.prototype.parse = function (xmlString, skipDeclaration) {
        if (skipDeclaration === void 0) { skipDeclaration = true; }
        if (skipDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");
        return new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
    };
    XmlParser.prototype.elements = function (elem, localName) {
        if (localName === void 0) { localName = null; }
        var result = [];
        for (var i = 0, l = elem.childNodes.length; i < l; i++) {
            var c = elem.childNodes.item(i);
            if (c.nodeType == 1 && (localName == null || c.localName == localName))
                result.push(c);
        }
        return result;
    };
    XmlParser.prototype.element = function (elem, localName) {
        for (var i = 0, l = elem.childNodes.length; i < l; i++) {
            var c = elem.childNodes.item(i);
            if (c.nodeType == 1 && c.localName == localName)
                return c;
        }
        return null;
    };
    XmlParser.prototype.elementAttr = function (elem, localName, attrLocalName) {
        var el = this.element(elem, localName);
        return el ? this.attr(el, attrLocalName) : undefined;
    };
    XmlParser.prototype.attr = function (elem, localName) {
        for (var i = 0, l = elem.attributes.length; i < l; i++) {
            var a = elem.attributes.item(i);
            if (a.localName == localName)
                return a.value;
        }
        return null;
    };
    XmlParser.prototype.intAttr = function (node, attrName, defaultValue) {
        if (defaultValue === void 0) { defaultValue = null; }
        var val = this.attr(node, attrName);
        return val ? parseInt(val) : defaultValue;
    };
    XmlParser.prototype.floatAttr = function (node, attrName, defaultValue) {
        if (defaultValue === void 0) { defaultValue = null; }
        var val = this.attr(node, attrName);
        return val ? parseFloat(val) : defaultValue;
    };
    XmlParser.prototype.boolAttr = function (node, attrName, defaultValue) {
        if (defaultValue === void 0) { defaultValue = null; }
        var v = this.attr(node, attrName);
        switch (v) {
            case "true": return true;
            case "false": return false;
            case "1": return true;
            case "0": return false;
            default: return defaultValue;
        }
    };
    XmlParser.prototype.lengthAttr = function (node, attrName, usage) {
        if (usage === void 0) { usage = common_1.LengthUsage.Dxa; }
        return (0, common_1.convertLength)(this.attr(node, attrName), usage);
    };
    return XmlParser;
}());
exports.XmlParser = XmlParser;
var globalXmlParser = new XmlParser();
exports["default"] = globalXmlParser;


/***/ }),

/***/ "./src/styles/styles-part.ts":
/*!***********************************!*\
  !*** ./src/styles/styles-part.ts ***!
  \***********************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.StylesPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var StylesPart = (function (_super) {
    __extends(StylesPart, _super);
    function StylesPart(pkg, path, parser) {
        var _this = _super.call(this, pkg, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    StylesPart.prototype.parseXml = function (root) {
        this.styles = this._documentParser.parseStylesFile(root);
    };
    return StylesPart;
}(part_1.Part));
exports.StylesPart = StylesPart;


/***/ }),

/***/ "./src/theme/theme-part.ts":
/*!*********************************!*\
  !*** ./src/theme/theme-part.ts ***!
  \*********************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.ThemePart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var theme_1 = __webpack_require__(/*! ./theme */ "./src/theme/theme.ts");
var ThemePart = (function (_super) {
    __extends(ThemePart, _super);
    function ThemePart(pkg, path) {
        return _super.call(this, pkg, path) || this;
    }
    ThemePart.prototype.parseXml = function (root) {
        this.theme = (0, theme_1.parseTheme)(root, this._package.xmlParser);
    };
    return ThemePart;
}(part_1.Part));
exports.ThemePart = ThemePart;


/***/ }),

/***/ "./src/theme/theme.ts":
/*!****************************!*\
  !*** ./src/theme/theme.ts ***!
  \****************************/
/***/ ((__unused_webpack_module, exports) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.parseFontInfo = exports.parseFontScheme = exports.parseColorScheme = exports.parseTheme = exports.DmlTheme = void 0;
var DmlTheme = (function () {
    function DmlTheme() {
    }
    return DmlTheme;
}());
exports.DmlTheme = DmlTheme;
function parseTheme(elem, xml) {
    var result = new DmlTheme();
    var themeElements = xml.element(elem, "themeElements");
    for (var _i = 0, _a = xml.elements(themeElements); _i < _a.length; _i++) {
        var el = _a[_i];
        switch (el.localName) {
            case "clrScheme":
                result.colorScheme = parseColorScheme(el, xml);
                break;
            case "fontScheme":
                result.fontScheme = parseFontScheme(el, xml);
                break;
        }
    }
    return result;
}
exports.parseTheme = parseTheme;
function parseColorScheme(elem, xml) {
    var result = {
        name: xml.attr(elem, "name"),
        colors: {}
    };
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var el = _a[_i];
        var srgbClr = xml.element(el, "srgbClr");
        var sysClr = xml.element(el, "sysClr");
        if (srgbClr) {
            result.colors[el.localName] = xml.attr(srgbClr, "val");
        }
        else if (sysClr) {
            result.colors[el.localName] = xml.attr(sysClr, "lastClr");
        }
    }
    return result;
}
exports.parseColorScheme = parseColorScheme;
function parseFontScheme(elem, xml) {
    var result = {
        name: xml.attr(elem, "name"),
    };
    for (var _i = 0, _a = xml.elements(elem); _i < _a.length; _i++) {
        var el = _a[_i];
        switch (el.localName) {
            case "majorFont":
                result.majorFont = parseFontInfo(el, xml);
                break;
            case "minorFont":
                result.minorFont = parseFontInfo(el, xml);
                break;
        }
    }
    return result;
}
exports.parseFontScheme = parseFontScheme;
function parseFontInfo(elem, xml) {
    return {
        latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
        eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
        csTypeface: xml.elementAttr(elem, "cs", "typeface"),
    };
}
exports.parseFontInfo = parseFontInfo;


/***/ }),

/***/ "./src/utils.ts":
/*!**********************!*\
  !*** ./src/utils.ts ***!
  \**********************/
/***/ (function(__unused_webpack_module, exports) {


var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.mergeDeep = exports.isObject = exports.keyBy = exports.resolvePath = exports.splitPath = exports.appendClass = exports.addElementClass = void 0;
function addElementClass(element, className) {
    return element.className = appendClass(element.className, className);
}
exports.addElementClass = addElementClass;
function appendClass(classList, className) {
    return (!classList) ? className : "".concat(classList, " ").concat(className);
}
exports.appendClass = appendClass;
function splitPath(path) {
    var si = path.lastIndexOf('/') + 1;
    var folder = si == 0 ? "" : path.substring(0, si);
    var fileName = si == 0 ? path : path.substring(si);
    return [folder, fileName];
}
exports.splitPath = splitPath;
function resolvePath(path, base) {
    try {
        var prefix = "file://docx/";
        var url = new URL(path, prefix + base).toString();
        return url.substr(prefix.length);
    }
    catch (_a) {
        return "".concat(base).concat(path);
    }
}
exports.resolvePath = resolvePath;
function keyBy(array, by) {
    return array.reduce(function (a, x) {
        a[by(x)] = x;
        return a;
    }, {});
}
exports.keyBy = keyBy;
function isObject(item) {
    return (item && typeof item === 'object' && !Array.isArray(item));
}
exports.isObject = isObject;
function mergeDeep(target) {
    var _a;
    var sources = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        sources[_i - 1] = arguments[_i];
    }
    if (!sources.length)
        return target;
    var source = sources.shift();
    if (isObject(target) && isObject(source)) {
        for (var key in source) {
            if (isObject(source[key])) {
                var val = (_a = target[key]) !== null && _a !== void 0 ? _a : (target[key] = {});
                mergeDeep(val, source[key]);
            }
            else {
                target[key] = source[key];
            }
        }
    }
    return mergeDeep.apply(void 0, __spreadArray([target], sources, false));
}
exports.mergeDeep = mergeDeep;


/***/ }),

/***/ "./src/word-document.ts":
/*!******************************!*\
  !*** ./src/word-document.ts ***!
  \******************************/
/***/ ((__unused_webpack_module, exports, __webpack_require__) => {


Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.deobfuscate = exports.WordDocument = void 0;
var relationship_1 = __webpack_require__(/*! ./common/relationship */ "./src/common/relationship.ts");
var font_table_1 = __webpack_require__(/*! ./font-table/font-table */ "./src/font-table/font-table.ts");
var open_xml_package_1 = __webpack_require__(/*! ./common/open-xml-package */ "./src/common/open-xml-package.ts");
var document_part_1 = __webpack_require__(/*! ./document/document-part */ "./src/document/document-part.ts");
var utils_1 = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var numbering_part_1 = __webpack_require__(/*! ./numbering/numbering-part */ "./src/numbering/numbering-part.ts");
var styles_part_1 = __webpack_require__(/*! ./styles/styles-part */ "./src/styles/styles-part.ts");
var footer_part_1 = __webpack_require__(/*! ./footer/footer-part */ "./src/footer/footer-part.ts");
var header_part_1 = __webpack_require__(/*! ./header/header-part */ "./src/header/header-part.ts");
var extended_props_part_1 = __webpack_require__(/*! ./document-props/extended-props-part */ "./src/document-props/extended-props-part.ts");
var core_props_part_1 = __webpack_require__(/*! ./document-props/core-props-part */ "./src/document-props/core-props-part.ts");
var theme_part_1 = __webpack_require__(/*! ./theme/theme-part */ "./src/theme/theme-part.ts");
var footnotes_part_1 = __webpack_require__(/*! ./footnotes/footnotes-part */ "./src/footnotes/footnotes-part.ts");
var topLevelRels = [
    { type: relationship_1.RelationshipTypes.OfficeDocument, target: "word/document.xml" },
    { type: relationship_1.RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
    { type: relationship_1.RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
];
var WordDocument = (function () {
    function WordDocument() {
        this.parts = [];
        this.partsMap = {};
    }
    WordDocument.load = function (blob, parser, options) {
        var d = new WordDocument();
        d._parser = parser;
        return open_xml_package_1.OpenXmlPackage.load(blob, options)
            .then(function (pkg) {
            d._package = pkg;
            return d._package.loadRelationships();
        }).then(function (rels) {
            d.rels = rels;
            var tasks = topLevelRels.map(function (rel) {
                var _a;
                var r = (_a = rels.find(function (x) { return x.type === rel.type; })) !== null && _a !== void 0 ? _a : rel;
                return d.loadRelationshipPart(r.target, r.type);
            });
            return Promise.all(tasks);
        }).then(function () { return d; });
    };
    WordDocument.prototype.save = function (type) {
        if (type === void 0) { type = "blob"; }
        return this._package.save(type);
    };
    WordDocument.prototype.loadRelationshipPart = function (path, type) {
        var _this = this;
        if (this.partsMap[path])
            return Promise.resolve(this.partsMap[path]);
        if (!this._package.get(path))
            return Promise.resolve(null);
        var part = null;
        switch (type) {
            case relationship_1.RelationshipTypes.OfficeDocument:
                this.documentPart = part = new document_part_1.DocumentPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.FontTable:
                this.fontTablePart = part = new font_table_1.FontTablePart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.Numbering:
                this.numberingPart = part = new numbering_part_1.NumberingPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Styles:
                this.stylesPart = part = new styles_part_1.StylesPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Theme:
                part = new theme_part_1.ThemePart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.Footnotes:
                this.footnotesPart = part = new footnotes_part_1.FootnotesPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Footer:
                part = new footer_part_1.FooterPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Header:
                part = new header_part_1.HeaderPart(this._package, path, this._parser);
                break;
            case relationship_1.RelationshipTypes.CoreProperties:
                this.corePropsPart = part = new core_props_part_1.CorePropsPart(this._package, path);
                break;
            case relationship_1.RelationshipTypes.ExtendedProperties:
                this.extendedPropsPart = part = new extended_props_part_1.ExtendedPropsPart(this._package, path);
                break;
        }
        if (part == null)
            return Promise.resolve(null);
        this.partsMap[path] = part;
        this.parts.push(part);
        return part.load().then(function () {
            if (part.rels == null || part.rels.length == 0)
                return part;
            var folder = (0, utils_1.splitPath)(part.path)[0];
            var rels = part.rels.map(function (rel) {
                return _this.loadRelationshipPart((0, utils_1.resolvePath)(rel.target, folder), rel.type);
            });
            return Promise.all(rels).then(function () { return part; });
        });
    };
    WordDocument.prototype.loadDocumentImage = function (id) {
        return this.loadResource(this.documentPart, id, "blob")
            .then(function (x) { return x ? URL.createObjectURL(x) : null; });
    };
    WordDocument.prototype.loadNumberingImage = function (id) {
        return this.loadResource(this.numberingPart, id, "blob")
            .then(function (x) { return x ? URL.createObjectURL(x) : null; });
    };
    WordDocument.prototype.loadFont = function (id, key) {
        return this.loadResource(this.fontTablePart, id, "uint8array")
            .then(function (x) { return x ? URL.createObjectURL(new Blob([deobfuscate(x, key)])) : x; });
    };
    WordDocument.prototype.findPartByRelId = function (id, basePart) {
        var _a;
        if (basePart === void 0) { basePart = null; }
        var rel = ((_a = basePart.rels) !== null && _a !== void 0 ? _a : this.rels).find(function (r) { return r.id == id; });
        var folder = basePart ? (0, utils_1.splitPath)(basePart.path)[0] : '';
        return rel ? this.partsMap[(0, utils_1.resolvePath)(rel.target, folder)] : null;
    };
    WordDocument.prototype.getPathById = function (part, id) {
        var rel = part.rels.find(function (x) { return x.id == id; });
        var folder = (0, utils_1.splitPath)(part.path)[0];
        return rel ? (0, utils_1.resolvePath)(rel.target, folder) : null;
    };
    WordDocument.prototype.loadResource = function (part, id, outputType) {
        var path = this.getPathById(part, id);
        return path ? this._package.load(path, outputType) : Promise.resolve(null);
    };
    return WordDocument;
}());
exports.WordDocument = WordDocument;
function deobfuscate(data, guidKey) {
    var len = 16;
    var trimmed = guidKey.replace(/{|}|-/g, "");
    var numbers = new Array(len);
    for (var i = 0; i < len; i++)
        numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);
    for (var i = 0; i < 32; i++)
        data[i] = data[i] ^ numbers[i % len];
    return data;
}
exports.deobfuscate = deobfuscate;


/***/ }),

/***/ "jszip":
/*!************************!*\
  !*** external "JSZip" ***!
  \************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE_jszip__;

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	var __webpack_exports__ = __webpack_require__("./src/docx-preview.ts");
/******/ 	
/******/ 	return __webpack_exports__;
/******/ })()
;
});
//# sourceMappingURL=docx-preview.js.map