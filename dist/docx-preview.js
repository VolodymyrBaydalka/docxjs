(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory(require("JSZip"));
	else if(typeof define === 'function' && define.amd)
		define(["JSZip"], factory);
	else if(typeof exports === 'object')
		exports["docx"] = factory(require("JSZip"));
	else
		root["docx"] = factory(root["JSZip"]);
})(window, function(__WEBPACK_EXTERNAL_MODULE_jszip__) {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/docx-preview.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/common/package.ts":
/*!*******************************!*\
  !*** ./src/common/package.ts ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.Package = void 0;
var xml_parser_1 = __webpack_require__(/*! ../parser/xml-parser */ "./src/parser/xml-parser.ts");
var utils_1 = __webpack_require__(/*! ../utils */ "./src/utils.ts");
var relationship_1 = __webpack_require__(/*! ./relationship */ "./src/common/relationship.ts");
var Package = (function () {
    function Package(_zip) {
        this._zip = _zip;
        this.xmlParser = new xml_parser_1.XmlParser();
    }
    Package.prototype.exists = function (path) {
        return this._zip.files[path] != null;
    };
    Package.prototype.load = function (path, type) {
        var _this = this;
        var file = this._zip.files[path];
        if (file == null)
            return Promise.resolve(null);
        if (type == "xml")
            return file.async("string").then(function (t) { return _this.xmlParser.parse(t); });
        return file.async(type);
    };
    Package.prototype.loadRelationships = function (path) {
        var _this = this;
        if (path === void 0) { path = null; }
        var relsPath = "_rels/.rels";
        if (path != null) {
            var _a = utils_1.splitPath(path), f = _a[0], fn = _a[1];
            relsPath = f + "_rels/" + fn + ".rels";
        }
        return this.load(relsPath, "xml").then(function (xml) {
            return xml == null ? null : relationship_1.parseRelationships(xml, _this.xmlParser);
        });
    };
    return Package;
}());
exports.Package = Package;


/***/ }),

/***/ "./src/common/part.ts":
/*!****************************!*\
  !*** ./src/common/part.ts ***!
  \****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.Part = void 0;
var Part = (function () {
    function Part(path) {
        this.path = path;
    }
    Part.prototype.load = function (pkg) {
        var _this = this;
        return pkg.loadRelationships(this.path).then(function (rels) {
            _this.rels = rels;
        });
    };
    return Part;
}());
exports.Part = Part;


/***/ }),

/***/ "./src/common/relationship.ts":
/*!************************************!*\
  !*** ./src/common/relationship.ts ***!
  \************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
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
})(RelationshipTypes = exports.RelationshipTypes || (exports.RelationshipTypes = {}));
function parseRelationships(root, xmlParser) {
    return xmlParser.elements(root).map(function (e) { return ({
        id: xmlParser.attr(e, "Id"),
        type: xmlParser.attr(e, "Type"),
        target: xmlParser.attr(e, "Target"),
        targetMode: xmlParser.attr(e, "TargetMode")
    }); });
}
exports.parseRelationships = parseRelationships;


/***/ }),

/***/ "./src/document-parser.ts":
/*!********************************!*\
  !*** ./src/document-parser.ts ***!
  \********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.DocumentParser = exports.autos = void 0;
var dom_1 = __webpack_require__(/*! ./dom/dom */ "./src/dom/dom.ts");
var utils = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var common_1 = __webpack_require__(/*! ./dom/common */ "./src/dom/common.ts");
var common_2 = __webpack_require__(/*! ./parser/common */ "./src/parser/common.ts");
var paragraph_1 = __webpack_require__(/*! ./parser/paragraph */ "./src/parser/paragraph.ts");
var section_1 = __webpack_require__(/*! ./parser/section */ "./src/parser/section.ts");
exports.autos = {
    shd: "white",
    color: "black",
    highlight: "transparent"
};
var DocumentParser = (function () {
    function DocumentParser() {
        this.skipDeclaration = true;
        this.ignoreWidth = false;
        this.debug = false;
    }
    DocumentParser.prototype.parseDocumentFile = function (xmlString) {
        var _this = this;
        var result = {
            type: dom_1.DomType.Document,
            children: [],
            style: {},
            props: null
        };
        var xbody = xml.byTagName(xml.parse(xmlString, this.skipDeclaration), "body");
        xml.foreach(xbody, function (elem) {
            switch (elem.localName) {
                case "p":
                    result.children.push(_this.parseParagraph(elem));
                    break;
                case "tbl":
                    result.children.push(_this.parseTable(elem));
                    break;
                case "sectPr":
                    result.props = section_1.parseSectionProperties(elem);
                    break;
            }
        });
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
                    var rPr = xml.byTagName(c, "rPr");
                    if (rPr)
                        result.styles.push({
                            target: "span",
                            values: _this.parseDefaultProperties(rPr, {})
                        });
                    break;
                case "pPrDefault":
                    var pPr = xml.byTagName(c, "pPr");
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
    DocumentParser.prototype.parseCommonProperties = function (elem, props) {
        if (elem.namespaceURI != common_1.ns.wordml)
            return;
        switch (elem.localName) {
            case "color":
                props.color = common_2.colorAttr(elem, elem.namespaceURI, "val");
                break;
            case "sz":
                props.fontSize = common_2.lengthAttr(elem, elem.namespaceURI, "val", common_2.LengthUsage.FontSize);
                break;
        }
    };
    DocumentParser.prototype.parseStyle = function (node) {
        var _this = this;
        var result = {
            id: xml.className(node, "styleId"),
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
                case "aliases":
                    result.aliases = xml.stringAttr(n, "val").split(",");
                    break;
                case "pPr":
                    result.styles.push({
                        target: "p",
                        values: _this.parseDefaultProperties(n, {})
                    });
                    break;
                case "rPr":
                    result.styles.push({
                        target: "span",
                        values: _this.parseDefaultProperties(n, {})
                    });
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
                    _this.debug && console.warn("DOCX: Unknown style element: " + n.localName);
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
    DocumentParser.prototype.parseNumberingFile = function (xmlString) {
        var _this = this;
        var result = [];
        var xnums = xml.parse(xmlString, this.skipDeclaration);
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
        var pict = xml.byTagName(elem, "pict");
        var shape = pict && xml.byTagName(pict, "shape");
        var imagedata = shape && xml.byTagName(shape, "imagedata");
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
            style: {}
        };
        xml.foreach(node, function (n) {
            switch (n.localName) {
                case "pPr":
                    _this.parseDefaultProperties(n, result.style);
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
            }
        });
        return result;
    };
    DocumentParser.prototype.parseParagraph = function (node) {
        var _this = this;
        var result = { type: dom_1.DomType.Paragraph, children: [], props: {} };
        xml.foreach(node, function (c) {
            switch (c.localName) {
                case "r":
                    result.children.push(_this.parseRun(c, result));
                    break;
                case "hyperlink":
                    result.children.push(_this.parseHyperlink(c, result));
                    break;
                case "bookmarkStart":
                    result.children.push(_this.parseBookmark(c));
                    break;
                case "pPr":
                    _this.parseParagraphProperties(c, result);
                    _this.parseCommonProperties(c, result.props);
                    break;
            }
        });
        return result;
    };
    DocumentParser.prototype.parseParagraphProperties = function (elem, paragraph) {
        var _this = this;
        this.parseDefaultProperties(elem, paragraph.style = {}, null, function (c) {
            if (paragraph_1.parseParagraphProperties(c, paragraph.props))
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
            paragraph.style["float"] = "left";
    };
    DocumentParser.prototype.parseBookmark = function (node) {
        var result = { type: dom_1.DomType.Run };
        result.id = xml.stringAttr(node, "name");
        return result;
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
                        break: "page"
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
        this.parseDefaultProperties(elem, run.style = {}, null, function (c) {
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
        for (var _i = 0, _a = xml.elements(node); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "inline":
                case "anchor":
                    return this.parseDrawingWrapper(n);
            }
        }
    };
    DocumentParser.prototype.parseDrawingWrapper = function (node) {
        var result = { type: dom_1.DomType.Drawing, children: [], style: {} };
        var isAnchor = node.localName == "anchor";
        var wrapType = null;
        var simplePos = xml.boolAttr(node, "simplePos");
        var posX = { relative: "page", align: "left", offset: "0" };
        var posY = { relative: "page", align: "top", offset: "0" };
        for (var _i = 0, _a = xml.elements(node); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "simplePos":
                    if (simplePos) {
                        posX.offset = xml.sizeAttr(n, "x", SizeType.Emu);
                        posY.offset = xml.sizeAttr(n, "y", SizeType.Emu);
                    }
                    break;
                case "extent":
                    result.style["width"] = xml.sizeAttr(n, "cx", SizeType.Emu);
                    result.style["height"] = xml.sizeAttr(n, "cy", SizeType.Emu);
                    break;
                case "positionH":
                case "positionV":
                    if (!simplePos) {
                        var pos = n.localName == "positionH" ? posX : posY;
                        var alignNode = xml.byTagName(n, "align");
                        var offsetNode = xml.byTagName(n, "posOffset");
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
            result.style['display'] = 'block';
            if (posX.align) {
                result.style['text-align'] = posX.align;
                result.style['width'] = "100%";
            }
        }
        else if (wrapType == "wrapNone") {
            result.style['display'] = 'block';
            result.style['position'] = 'relative';
            result.style["width"] = "0px";
            result.style["height"] = "0px";
            if (posX.offset)
                result.style["left"] = posX.offset;
            if (posY.offset)
                result.style["top"] = posY.offset;
        }
        else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
            result.style["float"] = posX.align;
        }
        return result;
    };
    DocumentParser.prototype.parseGraphic = function (elem) {
        var graphicData = xml.byTagName(elem, "graphicData");
        for (var _i = 0, _a = xml.elements(graphicData); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "pic":
                    return this.parsePicture(n);
            }
        }
        return null;
    };
    DocumentParser.prototype.parsePicture = function (elem) {
        var result = { type: dom_1.DomType.Image, src: "", style: {} };
        var blipFill = xml.byTagName(elem, "blipFill");
        var blip = xml.byTagName(blipFill, "blip");
        result.src = xml.stringAttr(blip, "embed");
        var spPr = xml.byTagName(elem, "spPr");
        var xfrm = xml.byTagName(spPr, "xfrm");
        result.style["position"] = "relative";
        for (var _i = 0, _a = xml.elements(xfrm); _i < _a.length; _i++) {
            var n = _a[_i];
            switch (n.localName) {
                case "ext":
                    result.style["width"] = xml.sizeAttr(n, "cx", SizeType.Emu);
                    result.style["height"] = xml.sizeAttr(n, "cy", SizeType.Emu);
                    break;
                case "off":
                    result.style["left"] = xml.sizeAttr(n, "x", SizeType.Emu);
                    result.style["top"] = xml.sizeAttr(n, "y", SizeType.Emu);
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
        table.style = {};
        table.cellStyle = {};
        this.parseDefaultProperties(elem, table.style, table.cellStyle, function (c) {
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
        switch (table.style["text-align"]) {
            case "center":
                delete table.style["text-align"];
                table.style["margin-left"] = "auto";
                table.style["margin-right"] = "auto";
                break;
            case "right":
                delete table.style["text-align"];
                table.style["margin-left"] = "auto";
                break;
        }
    };
    DocumentParser.prototype.parseTablePosition = function (node, table) {
        var topFromText = xml.sizeAttr(node, "topFromText");
        var bottomFromText = xml.sizeAttr(node, "bottomFromText");
        var rightFromText = xml.sizeAttr(node, "rightFromText");
        var leftFromText = xml.sizeAttr(node, "leftFromText");
        table.style["float"] = 'left';
        table.style["margin-bottom"] = values.addSize(table.style["margin-bottom"], bottomFromText);
        table.style["margin-left"] = values.addSize(table.style["margin-left"], leftFromText);
        table.style["margin-right"] = values.addSize(table.style["margin-right"], rightFromText);
        table.style["margin-top"] = values.addSize(table.style["margin-top"], topFromText);
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
        row.style = this.parseDefaultProperties(elem, {}, null, function (c) {
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
        cell.style = this.parseDefaultProperties(elem, {}, null, function (c) {
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
        var spacing = null;
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
                    if (_this.ignoreWidth)
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
                    style["font-weight"] = values.valueOfBold(c);
                    break;
                case "i":
                    style["font-style"] = "italic";
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
                    style["vertical-align"] = xml.stringAttr(c, "val");
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
                        _this.debug && console.warn("DOCX: Unknown document element: " + c.localName);
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
        var left = xml.sizeAttr(node, "left");
        var start = xml.sizeAttr(node, "start");
        var right = xml.sizeAttr(node, "right");
        var end = xml.sizeAttr(node, "end");
        if (firstLine)
            style["text-indent"] = firstLine;
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
                    style["line-height"] = "" + (line / 240).toFixed(2);
                    break;
                case "atLeast":
                    style["line-height"] = "calc(100% + " + line / 20 + "pt)";
                    break;
                default:
                    style["line-height"] = style["min-height"] = line / 20 + "pt";
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
    xml.parse = function (xmlString, skipDeclaration) {
        if (skipDeclaration === void 0) { skipDeclaration = true; }
        if (skipDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");
        return new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
    };
    xml.elements = function (node, tagName) {
        if (tagName === void 0) { tagName = null; }
        var result = [];
        for (var i = 0; i < node.childNodes.length; i++) {
            var n = node.childNodes[i];
            if (n.nodeType == 1 && (tagName == null || n.localName == tagName))
                result.push(n);
        }
        return result;
    };
    xml.foreach = function (node, cb) {
        for (var i = 0; i < node.childNodes.length; i++) {
            var n = node.childNodes[i];
            if (n.nodeType == 1)
                cb(n);
        }
    };
    xml.byTagName = function (elem, tagName) {
        for (var i = 0; i < elem.childNodes.length; i++) {
            var n = elem.childNodes[i];
            if (n.nodeType == 1 && n.localName == tagName)
                return elem.childNodes[i];
        }
        return null;
    };
    xml.elementStringAttr = function (elem, nodeName, attrName) {
        var n = xml.byTagName(elem, nodeName);
        return n ? xml.stringAttr(n, attrName) : null;
    };
    xml.stringAttr = function (node, attrName) {
        var elem = node;
        for (var i = 0; i < elem.attributes.length; i++) {
            var attr = elem.attributes.item(i);
            if (attr.localName == attrName)
                return attr.value;
        }
        return null;
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
        return v ? "#" + v : defValue;
    };
    xml.boolAttr = function (node, attrName, defValue) {
        if (defValue === void 0) { defValue = false; }
        var v = xml.stringAttr(node, attrName);
        switch (v) {
            case "1": return true;
            case "0": return false;
        }
        return defValue;
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
        if (val == null || val.indexOf("pt") > -1)
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
        return size + " solid " + (color == "auto" ? "black" : color);
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
        return "calc(" + a + " + " + b + ")";
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

/***/ "./src/docx-preview.ts":
/*!*****************************!*\
  !*** ./src/docx-preview.ts ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

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
Object.defineProperty(exports, "__esModule", { value: true });
exports.renderAsync = void 0;
var word_document_1 = __webpack_require__(/*! ./word-document */ "./src/word-document.ts");
var document_parser_1 = __webpack_require__(/*! ./document-parser */ "./src/document-parser.ts");
var html_renderer_1 = __webpack_require__(/*! ./html-renderer */ "./src/html-renderer.ts");
function renderAsync(data, bodyContainer, styleContainer, userOptions) {
    if (styleContainer === void 0) { styleContainer = null; }
    if (userOptions === void 0) { userOptions = null; }
    var parser = new document_parser_1.DocumentParser();
    var renderer = new html_renderer_1.HtmlRenderer(window.document);
    var options = __assign({ ignoreHeight: false, ignoreWidth: false, ignoreFonts: false, breakPages: true, debug: false, experimental: false, className: "docx", inWrapper: true }, userOptions);
    parser.ignoreWidth = options.ignoreWidth;
    parser.debug = options.debug || parser.debug;
    renderer.className = options.className || "docx";
    renderer.inWrapper = options.inWrapper;
    return word_document_1.WordDocument.load(data, parser).then(function (doc) {
        renderer.render(doc, bodyContainer, styleContainer, options);
        return doc;
    });
}
exports.renderAsync = renderAsync;


/***/ }),

/***/ "./src/dom/common.ts":
/*!***************************!*\
  !*** ./src/dom/common.ts ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.ns = void 0;
exports.ns = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
};


/***/ }),

/***/ "./src/dom/document-part.ts":
/*!**********************************!*\
  !*** ./src/dom/document-part.ts ***!
  \**********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.DocumentPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var DocumentPart = (function (_super) {
    __extends(DocumentPart, _super);
    function DocumentPart(path, parser) {
        var _this = _super.call(this, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    DocumentPart.prototype.load = function (pkg) {
        var _this = this;
        return _super.prototype.load.call(this, pkg)
            .then(function () { return pkg.load(_this.path, "string"); })
            .then(function (xml) {
            _this.body = _this._documentParser.parseDocumentFile(xml);
        });
    };
    return DocumentPart;
}(part_1.Part));
exports.DocumentPart = DocumentPart;


/***/ }),

/***/ "./src/dom/dom.ts":
/*!************************!*\
  !*** ./src/dom/dom.ts ***!
  \************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
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
})(DomType = exports.DomType || (exports.DomType = {}));


/***/ }),

/***/ "./src/font-table/font-table.ts":
/*!**************************************!*\
  !*** ./src/font-table/font-table.ts ***!
  \**************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.FontTablePart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var fonts_1 = __webpack_require__(/*! ./fonts */ "./src/font-table/fonts.ts");
var FontTablePart = (function (_super) {
    __extends(FontTablePart, _super);
    function FontTablePart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    FontTablePart.prototype.load = function (pkg) {
        var _this = this;
        return _super.prototype.load.call(this, pkg)
            .then(function () { return pkg.load(_this.path, "xml"); })
            .then(function (el) {
            _this.fonts = fonts_1.parseFonts(el, pkg.xmlParser);
        });
    };
    return FontTablePart;
}(part_1.Part));
exports.FontTablePart = FontTablePart;


/***/ }),

/***/ "./src/font-table/fonts.ts":
/*!*********************************!*\
  !*** ./src/font-table/fonts.ts ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.parseFonts = void 0;
function parseFonts(root, xmlParser) {
    var result = [];
    for (var _i = 0, _a = xmlParser.elements(root); _i < _a.length; _i++) {
        var el = _a[_i];
        var font = {
            name: xmlParser.attr(el, "name")
        };
        var embed = xmlParser.element(el, "embedRegular");
        if (embed) {
            font.fontKey = xmlParser.attr(embed, "fontKey");
            font.refId = xmlParser.attr(embed, "id");
        }
        result.push(font);
    }
    return result;
}
exports.parseFonts = parseFonts;


/***/ }),

/***/ "./src/html-renderer.ts":
/*!******************************!*\
  !*** ./src/html-renderer.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

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
Object.defineProperty(exports, "__esModule", { value: true });
exports.HtmlRenderer = void 0;
var dom_1 = __webpack_require__(/*! ./dom/dom */ "./src/dom/dom.ts");
var utils_1 = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var javascript_1 = __webpack_require__(/*! ./javascript */ "./src/javascript.ts");
var HtmlRenderer = (function () {
    function HtmlRenderer(htmlDocument) {
        this.htmlDocument = htmlDocument;
        this.inWrapper = true;
        this.className = "docx";
    }
    HtmlRenderer.prototype.render = function (document, bodyContainer, styleContainer, options) {
        if (styleContainer === void 0) { styleContainer = null; }
        this.document = document;
        this.options = options;
        styleContainer = styleContainer || bodyContainer;
        removeAllElements(styleContainer);
        removeAllElements(bodyContainer);
        appendComment(styleContainer, "docxjs library predefined styles");
        styleContainer.appendChild(this.renderDefaultStyle());
        if (document.stylesPart != null) {
            appendComment(styleContainer, "docx document styles");
            styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
        }
        if (document.numberingPart) {
            appendComment(styleContainer, "docx document numbering styles");
            styleContainer.appendChild(this.renderNumbering(document.numberingPart.numberings, styleContainer));
        }
        if (!options.ignoreFonts && document.fontTablePart)
            this.renderFontTable(document.fontTablePart, styleContainer);
        var sectionElements = this.renderSections(document.documentPart.body);
        if (this.inWrapper) {
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
                var cssTest = "@font-face {\n                    font-family: \"" + f.name + "\";\n                    src: url(" + fontData + ");\n                }";
                appendComment(styleContainer, "Font " + f.name);
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
        return this.className + "_" + className;
    };
    HtmlRenderer.prototype.processStyles = function (styles) {
        var stylesMap = {};
        for (var _i = 0, _a = styles.filter(function (x) { return x.id != null; }); _i < _a.length; _i++) {
            var style = _a[_i];
            stylesMap[style.id] = style;
        }
        for (var _b = 0, _c = styles.filter(function (x) { return x.basedOn; }); _b < _c.length; _b++) {
            var style = _c[_b];
            var baseStyle = stylesMap[style.basedOn];
            if (baseStyle) {
                var _loop_2 = function (styleValues) {
                    baseValues = baseStyle.styles.filter(function (x) { return x.target == styleValues.target; });
                    if (baseValues && baseValues.length > 0)
                        this_2.copyStyleProperties(baseValues[0].values, styleValues.values);
                };
                var this_2 = this, baseValues;
                for (var _d = 0, _e = style.styles; _d < _e.length; _d++) {
                    var styleValues = _e[_d];
                    _loop_2(styleValues);
                }
            }
            else if (this.options.debug)
                console.warn("Can't find base style " + style.basedOn);
        }
        for (var _f = 0, styles_1 = styles; _f < styles_1.length; _f++) {
            var style = styles_1[_f];
            style.id = this.processClassName(style.id);
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
                c.style = this.copyStyleProperties(table.cellStyle, c.style, [
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
                elem.style.columnCount = "" + props.columns.numberOfColumns;
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
            var sectionElement = this.createSection(this.className, section.sectProps || document.props);
            this.renderElements(section.elements, document, sectionElement);
            result.push(sectionElement);
        }
        return result;
    };
    HtmlRenderer.prototype.splitBySection = function (elements) {
        var current = { sectProps: null, elements: [] };
        var result = [current];
        for (var _i = 0, elements_1 = elements; _i < elements_1.length; _i++) {
            var elem = elements_1[_i];
            current.elements.push(elem);
            if (elem.type == dom_1.DomType.Paragraph) {
                var p = elem;
                var sectProps = p.props.sectionProps;
                var pBreakIndex = -1;
                var rBreakIndex = -1;
                if (this.options.breakPages && p.children) {
                    pBreakIndex = p.children.findIndex(function (r) {
                        var _a, _b;
                        rBreakIndex = (_b = (_a = r.children) === null || _a === void 0 ? void 0 : _a.findIndex(function (t) { return t.break == "page"; })) !== null && _b !== void 0 ? _b : -1;
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
        return result;
    };
    HtmlRenderer.prototype.renderLength = function (l) {
        return !l ? null : "" + l.value + l.type;
    };
    HtmlRenderer.prototype.renderWrapper = function () {
        var wrapper = document.createElement("div");
        wrapper.className = this.className + "-wrapper";
        return wrapper;
    };
    HtmlRenderer.prototype.renderDefaultStyle = function () {
        var styleText = "." + this.className + "-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } \n                ." + this.className + "-wrapper section." + this.className + " { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }\n                ." + this.className + " { color: black; }\n                section." + this.className + " { box-sizing: border-box; }\n                ." + this.className + " table { border-collapse: collapse; }\n                ." + this.className + " table td, ." + this.className + " table th { vertical-align: top; }\n                ." + this.className + " p { margin: 0pt; }";
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderNumbering = function (styles, styleContainer) {
        var _this = this;
        var styleText = "";
        var rootCounters = [];
        var _loop_3 = function () {
            selector = "p." + this_3.numberingClass(num.id, num.level);
            listStyleType = "none";
            if (num.levelText && num.format == "decimal") {
                var counter = this_3.numberingCounter(num.id, num.level);
                if (num.level > 0) {
                    styleText += this_3.styleToString("p." + this_3.numberingClass(num.id, num.level - 1), {
                        "counter-reset": counter
                    });
                }
                else {
                    rootCounters.push(counter);
                }
                styleText += this_3.styleToString(selector + ":before", {
                    "content": this_3.levelTextToContent(num.levelText, num.id),
                    "counter-increment": counter
                });
                styleText += this_3.styleToString(selector, __assign({ "display": "list-item", "list-style-position": "inside", "list-style-type": "none" }, num.style));
            }
            else if (num.bullet) {
                var valiable_1 = ("--" + this_3.className + "-" + num.bullet.src).toLowerCase();
                styleText += this_3.styleToString(selector + ":before", {
                    "content": "' '",
                    "display": "inline-block",
                    "background": "var(" + valiable_1 + ")"
                }, num.bullet.style);
                this_3.document.loadNumberingImage(num.bullet.src).then(function (data) {
                    var text = "." + _this.className + "-wrapper { " + valiable_1 + ": url(" + data + ") }";
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else {
                listStyleType = this_3.numFormatToCssValue(num.format);
            }
            styleText += this_3.styleToString(selector, __assign({ "display": "list-item", "list-style-position": "inside", "list-style-type": listStyleType }, num.style));
        };
        var this_3 = this, selector, listStyleType;
        for (var _i = 0, styles_2 = styles; _i < styles_2.length; _i++) {
            var num = styles_2[_i];
            _loop_3();
        }
        if (rootCounters.length > 0) {
            styleText += this.styleToString("." + this.className + "-wrapper", {
                "counter-reset": rootCounters.join(" ")
            });
        }
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderStyles = function (styles) {
        var styleText = "";
        var stylesMap = this.processStyles(styles);
        for (var _i = 0, styles_3 = styles; _i < styles_3.length; _i++) {
            var style = styles_3[_i];
            var subStyles = style.styles;
            if (style.linked) {
                var linkedStyle = style.linked && stylesMap[style.linked];
                if (linkedStyle)
                    subStyles = subStyles.concat(linkedStyle.styles);
                else if (this.options.debug)
                    console.warn("Can't find linked style " + style.linked);
            }
            for (var _a = 0, subStyles_1 = subStyles; _a < subStyles_1.length; _a++) {
                var subStyle = subStyles_1[_a];
                var selector = "";
                if (style.target == subStyle.target)
                    selector += style.target + "." + style.id;
                else if (style.target)
                    selector += style.target + "." + style.id + " " + subStyle.target;
                else
                    selector += "." + style.id + " " + subStyle.target;
                if (style.isDefault && style.target)
                    selector = "." + this.className + " " + style.target + ", " + selector;
                styleText += this.styleToString(selector, subStyle.values);
            }
        }
        return createStyleElement(styleText);
    };
    HtmlRenderer.prototype.renderElement = function (elem, parent) {
        switch (elem.type) {
            case dom_1.DomType.Paragraph:
                return this.renderParagraph(elem);
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
        }
        return null;
    };
    HtmlRenderer.prototype.renderChildren = function (elem, into) {
        return this.renderElements(elem.children, elem, into);
    };
    HtmlRenderer.prototype.renderElements = function (elems, parent, into) {
        var _this = this;
        if (elems == null)
            return null;
        var result = elems.map(function (e) { return _this.renderElement(e, parent); }).filter(function (e) { return e != null; });
        if (into)
            for (var _i = 0, result_1 = result; _i < result_1.length; _i++) {
                var c = result_1[_i];
                into.appendChild(c);
            }
        return result;
    };
    HtmlRenderer.prototype.renderParagraph = function (elem) {
        var result = this.htmlDocument.createElement("p");
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.style, result);
        this.renderCommonProeprties(result, elem.props);
        if (elem.props.numbering) {
            var numberingClass = this.numberingClass(elem.props.numbering.id, elem.props.numbering.level);
            result.className = utils_1.appendClass(result.className, numberingClass);
        }
        return result;
    };
    HtmlRenderer.prototype.renderCommonProeprties = function (elem, props) {
        if (props == null)
            return;
        if (props.color) {
            elem.style.color = props.color;
        }
        if (props.fontSize) {
            elem.style.fontSize = this.renderLength(props.fontSize);
        }
    };
    HtmlRenderer.prototype.renderHyperlink = function (elem) {
        var result = this.htmlDocument.createElement("a");
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.style, result);
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
        this.renderStyleValues(elem.style, result);
        return result;
    };
    HtmlRenderer.prototype.renderImage = function (elem) {
        var result = this.htmlDocument.createElement("img");
        this.renderStyleValues(elem.style, result);
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
    HtmlRenderer.prototype.renderSymbol = function (elem) {
        var span = this.htmlDocument.createElement("span");
        span.style.fontFamily = elem.font;
        span.innerHTML = "&#x" + elem.char + ";";
        return span;
    };
    HtmlRenderer.prototype.renderTab = function (elem) {
        var tabSpan = this.htmlDocument.createElement("span");
        tabSpan.innerHTML = "&emsp;";
        if (this.options.experimental) {
            setTimeout(function () {
                var paragraph = findParent(elem, dom_1.DomType.Paragraph);
                if (paragraph.props.tabs == null)
                    return;
                paragraph.props.tabs.sort(function (a, b) { return a.position.value - b.position.value; });
                tabSpan.style.display = "inline-block";
                javascript_1.updateTabStop(tabSpan, paragraph.props.tabs);
            }, 0);
        }
        return tabSpan;
    };
    HtmlRenderer.prototype.renderRun = function (elem) {
        if (elem.break)
            return elem.break == "page" ? null : this.htmlDocument.createElement("br");
        if (elem.fldCharType || elem.instrText)
            return null;
        var result = this.htmlDocument.createElement("span");
        if (elem.id)
            result.id = elem.id;
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.style, result);
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
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.style, result);
        if (elem.columns)
            result.appendChild(this.renderTableColumns(elem.columns));
        return result;
    };
    HtmlRenderer.prototype.renderTableColumns = function (columns) {
        var result = this.htmlDocument.createElement("colGroup");
        for (var _i = 0, columns_1 = columns; _i < columns_1.length; _i++) {
            var col = columns_1[_i];
            var colElem = this.htmlDocument.createElement("col");
            if (col.width)
                colElem.style.width = col.width + "px";
            result.appendChild(colElem);
        }
        return result;
    };
    HtmlRenderer.prototype.renderTableRow = function (elem) {
        var result = this.htmlDocument.createElement("tr");
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.style, result);
        return result;
    };
    HtmlRenderer.prototype.renderTableCell = function (elem) {
        var result = this.htmlDocument.createElement("td");
        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.style, result);
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
        return this.className + "-num-" + id + "-" + lvl;
    };
    HtmlRenderer.prototype.styleToString = function (selectors, values, cssText) {
        if (cssText === void 0) { cssText = null; }
        var result = selectors + " {\r\n";
        for (var key in values) {
            result += "  " + key + ": " + values[key] + ";\r\n";
        }
        if (cssText)
            result += ";" + cssText;
        return result + "}\r\n";
    };
    HtmlRenderer.prototype.numberingCounter = function (id, lvl) {
        return this.className + "-num-" + id + "-" + lvl;
    };
    HtmlRenderer.prototype.levelTextToContent = function (text, id) {
        var _this = this;
        var result = text.replace(/%\d*/g, function (s) {
            var lvl = parseInt(s.substring(1), 10) - 1;
            return "\"counter(" + _this.numberingCounter(id, lvl) + ")\"";
        });
        return '"' + result + '"';
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
    styleElement.type = "text/css";
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
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
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
    elem.style.width = (tab.position.value - left) + "pt";
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
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.NumberingPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var NumberingPart = (function (_super) {
    __extends(NumberingPart, _super);
    function NumberingPart(path, parser) {
        var _this = _super.call(this, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    NumberingPart.prototype.load = function (pkg) {
        var _this = this;
        return _super.prototype.load.call(this, pkg)
            .then(function () { return pkg.load(_this.path, "string"); })
            .then(function (xml) {
            _this.numberings = _this._documentParser.parseNumberingFile(xml);
        });
    };
    return NumberingPart;
}(part_1.Part));
exports.NumberingPart = NumberingPart;


/***/ }),

/***/ "./src/parser/common.ts":
/*!******************************!*\
  !*** ./src/parser/common.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.parseBorders = exports.parseBorder = exports.lengthAttr = exports.LengthUsage = exports.boolAttr = exports.colorAttr = exports.intAttr = exports.stringAttr = exports.elements = void 0;
var common_1 = __webpack_require__(/*! ../dom/common */ "./src/dom/common.ts");
function elements(elem, namespaceURI, localName) {
    if (namespaceURI === void 0) { namespaceURI = null; }
    if (localName === void 0) { localName = null; }
    var result = [];
    for (var i = 0; i < elem.childNodes.length; i++) {
        var n = elem.childNodes[i];
        if (n.nodeType == 1
            && (namespaceURI == null || n.namespaceURI == namespaceURI)
            && (localName == null || n.localName == localName))
            result.push(n);
    }
    return result;
}
exports.elements = elements;
function stringAttr(elem, namespaceURI, name) {
    return elem.getAttributeNS(namespaceURI, name);
}
exports.stringAttr = stringAttr;
function intAttr(elem, namespaceURI, name) {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? parseInt(val) : null;
}
exports.intAttr = intAttr;
function colorAttr(elem, namespaceURI, name) {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? "#" + val : null;
}
exports.colorAttr = colorAttr;
function boolAttr(elem, namespaceURI, name, defaultValue) {
    if (defaultValue === void 0) { defaultValue = false; }
    var val = elem.getAttributeNS(namespaceURI, name);
    if (val == null)
        return defaultValue;
    return val === "true" || val === "1";
}
exports.boolAttr = boolAttr;
exports.LengthUsage = {
    Dxa: { mul: 0.05, unit: "pt" },
    Emu: { mul: 1 / 12700, unit: "pt" },
    FontSize: { mul: 0.5, unit: "pt" },
    Border: { mul: 0.125, unit: "pt" },
    Percent: { mul: 0.02, unit: "%" },
    LineHeight: { mul: 1 / 240, unit: null }
};
function lengthAttr(elem, namespaceURI, name, usage) {
    if (usage === void 0) { usage = exports.LengthUsage.Dxa; }
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? { value: parseInt(val) * usage.mul, type: usage.unit } : null;
}
exports.lengthAttr = lengthAttr;
function parseBorder(elem) {
    return {
        type: stringAttr(elem, common_1.ns.wordml, "val"),
        color: colorAttr(elem, common_1.ns.wordml, "color"),
        size: lengthAttr(elem, common_1.ns.wordml, "sz", exports.LengthUsage.Border)
    };
}
exports.parseBorder = parseBorder;
function parseBorders(elem) {
    var result = {};
    for (var _i = 0, _a = elements(elem, common_1.ns.wordml); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "left":
                result.left = parseBorder(e);
                break;
            case "top":
                result.top = parseBorder(e);
                break;
            case "right":
                result.right = parseBorder(e);
                break;
            case "botton":
                result.botton = parseBorder(e);
                break;
        }
    }
    return result;
}
exports.parseBorders = parseBorders;


/***/ }),

/***/ "./src/parser/paragraph.ts":
/*!*********************************!*\
  !*** ./src/parser/paragraph.ts ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.parseParagraphProperties = void 0;
var xml = __webpack_require__(/*! ./common */ "./src/parser/common.ts");
var common_1 = __webpack_require__(/*! ../dom/common */ "./src/dom/common.ts");
var section_1 = __webpack_require__(/*! ./section */ "./src/parser/section.ts");
function parseParagraphProperties(elem, props) {
    if (elem.namespaceURI != common_1.ns.wordml)
        return false;
    switch (elem.localName) {
        case "tabs":
            props.tabs = parseTabs(elem);
            break;
        case "sectPr":
            props.sectionProps = section_1.parseSectionProperties(elem);
            break;
        case "numPr":
            props.numbering = parseNumbering(elem);
            break;
        case "spacing":
            props.lineSpacing = parseLineSpacing(elem);
            return false;
            break;
        default:
            return false;
    }
    return true;
}
exports.parseParagraphProperties = parseParagraphProperties;
function parseTabs(elem) {
    return xml.elements(elem, common_1.ns.wordml, "tab")
        .map(function (e) { return ({
        position: xml.lengthAttr(e, common_1.ns.wordml, "pos"),
        leader: xml.stringAttr(e, common_1.ns.wordml, "leader"),
        style: xml.stringAttr(e, common_1.ns.wordml, "val")
    }); });
}
function parseNumbering(elem) {
    var result = {};
    for (var _i = 0, _a = xml.elements(elem, common_1.ns.wordml); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "numId":
                result.id = xml.stringAttr(e, common_1.ns.wordml, "val");
                break;
            case "ilvl":
                result.level = xml.intAttr(e, common_1.ns.wordml, "val");
                break;
        }
    }
    return result;
}
function parseLineSpacing(elem) {
    return {
        before: xml.lengthAttr(elem, common_1.ns.wordml, "before"),
        after: xml.lengthAttr(elem, common_1.ns.wordml, "after"),
        line: xml.intAttr(elem, common_1.ns.wordml, "line"),
        lineRule: xml.stringAttr(elem, common_1.ns.wordml, "lineRule")
    };
}


/***/ }),

/***/ "./src/parser/section.ts":
/*!*******************************!*\
  !*** ./src/parser/section.ts ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.parseSectionProperties = void 0;
var common_1 = __webpack_require__(/*! ../dom/common */ "./src/dom/common.ts");
var xml = __webpack_require__(/*! ./common */ "./src/parser/common.ts");
function parseSectionProperties(elem) {
    var section = {};
    for (var _i = 0, _a = xml.elements(elem, common_1.ns.wordml); _i < _a.length; _i++) {
        var e = _a[_i];
        switch (e.localName) {
            case "pgSz":
                section.pageSize = {
                    width: xml.lengthAttr(e, common_1.ns.wordml, "w"),
                    height: xml.lengthAttr(e, common_1.ns.wordml, "h"),
                    orientation: xml.stringAttr(e, common_1.ns.wordml, "orient")
                };
                break;
            case "pgMar":
                section.pageMargins = {
                    left: xml.lengthAttr(e, common_1.ns.wordml, "left"),
                    right: xml.lengthAttr(e, common_1.ns.wordml, "right"),
                    top: xml.lengthAttr(e, common_1.ns.wordml, "top"),
                    bottom: xml.lengthAttr(e, common_1.ns.wordml, "bottom"),
                    header: xml.lengthAttr(e, common_1.ns.wordml, "header"),
                    footer: xml.lengthAttr(e, common_1.ns.wordml, "footer"),
                    gutter: xml.lengthAttr(e, common_1.ns.wordml, "gutter"),
                };
                break;
            case "cols":
                section.columns = parseColumns(e);
                break;
        }
    }
    return section;
}
exports.parseSectionProperties = parseSectionProperties;
function parseColumns(elem) {
    return {
        numberOfColumns: xml.intAttr(elem, common_1.ns.wordml, "num"),
        space: xml.lengthAttr(elem, common_1.ns.wordml, "space"),
        separator: xml.boolAttr(elem, common_1.ns.wordml, "sep"),
        equalWidth: xml.boolAttr(elem, common_1.ns.wordml, "equalWidth", true),
        columns: xml.elements(elem, common_1.ns.wordml, "col")
            .map(function (e) { return ({
            width: xml.lengthAttr(e, common_1.ns.wordml, "w"),
            space: xml.lengthAttr(e, common_1.ns.wordml, "space")
        }); })
    };
}


/***/ }),

/***/ "./src/parser/xml-parser.ts":
/*!**********************************!*\
  !*** ./src/parser/xml-parser.ts ***!
  \**********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.XmlParser = void 0;
var XmlParser = (function () {
    function XmlParser() {
    }
    XmlParser.prototype.parse = function (xmlString, skipDeclaration) {
        if (skipDeclaration === void 0) { skipDeclaration = true; }
        if (skipDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");
        return new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
    };
    XmlParser.prototype.elements = function (elem) {
        var result = [];
        for (var i = 0, l = elem.childNodes.length; i < l; i++) {
            var c = elem.childNodes.item(i);
            if (c.nodeType == 1)
                result.push(c);
        }
        return result;
    };
    XmlParser.prototype.element = function (elem, localName) {
        for (var i = 0, l = elem.childNodes.length; i < l; i++) {
            var c = elem.childNodes.item(i);
            if (c.nodeType == 1 && c.nodeName == localName)
                return c;
        }
        return null;
    };
    XmlParser.prototype.attr = function (elem, localName) {
        for (var i = 0, l = elem.attributes.length; i < l; i++) {
            var a = elem.attributes.item(i);
            if (a.localName == localName)
                return a.value;
        }
        return null;
    };
    return XmlParser;
}());
exports.XmlParser = XmlParser;


/***/ }),

/***/ "./src/styles/styles-part.ts":
/*!***********************************!*\
  !*** ./src/styles/styles-part.ts ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.StylesPart = void 0;
var part_1 = __webpack_require__(/*! ../common/part */ "./src/common/part.ts");
var document_parser_1 = __webpack_require__(/*! ../document-parser */ "./src/document-parser.ts");
var StylesPart = (function (_super) {
    __extends(StylesPart, _super);
    function StylesPart(path, parser) {
        var _this = _super.call(this, path) || this;
        _this._documentParser = parser;
        return _this;
    }
    StylesPart.prototype.load = function (pkg) {
        var _this = this;
        return _super.prototype.load.call(this, pkg)
            .then(function () { return pkg.load(_this.path, "string"); })
            .then(function (xml) {
            _this.styles = new document_parser_1.DocumentParser().parseStylesFile(xml);
        });
    };
    return StylesPart;
}(part_1.Part));
exports.StylesPart = StylesPart;


/***/ }),

/***/ "./src/utils.ts":
/*!**********************!*\
  !*** ./src/utils.ts ***!
  \**********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.splitPath = exports.appendClass = exports.addElementClass = void 0;
function addElementClass(element, className) {
    return element.className = appendClass(element.className, className);
}
exports.addElementClass = addElementClass;
function appendClass(classList, className) {
    return (!classList) ? className : classList + " " + className;
}
exports.appendClass = appendClass;
function splitPath(path) {
    var si = path.lastIndexOf('/') + 1;
    var folder = si == 0 ? "" : path.substring(0, si);
    var fileName = si == 0 ? path : path.substring(si);
    return [folder, fileName];
}
exports.splitPath = splitPath;


/***/ }),

/***/ "./src/word-document.ts":
/*!******************************!*\
  !*** ./src/word-document.ts ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.deobfuscate = exports.WordDocument = void 0;
var JSZip = __webpack_require__(/*! jszip */ "jszip");
var relationship_1 = __webpack_require__(/*! ./common/relationship */ "./src/common/relationship.ts");
var font_table_1 = __webpack_require__(/*! ./font-table/font-table */ "./src/font-table/font-table.ts");
var package_1 = __webpack_require__(/*! ./common/package */ "./src/common/package.ts");
var document_part_1 = __webpack_require__(/*! ./dom/document-part */ "./src/dom/document-part.ts");
var utils_1 = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var numbering_part_1 = __webpack_require__(/*! ./numbering/numbering-part */ "./src/numbering/numbering-part.ts");
var styles_part_1 = __webpack_require__(/*! ./styles/styles-part */ "./src/styles/styles-part.ts");
var WordDocument = (function () {
    function WordDocument() {
        this.parts = [];
        this.partsMap = {};
    }
    WordDocument.load = function (blob, parser) {
        var d = new WordDocument();
        d._parser = parser;
        return JSZip.loadAsync(blob)
            .then(function (zip) {
            d._package = new package_1.Package(zip);
            return d._package.loadRelationships();
        }).then(function (rels) {
            var _a;
            d.rels = rels;
            var _b = (_a = rels.find(function (x) { return x.type == relationship_1.RelationshipTypes.OfficeDocument; })) !== null && _a !== void 0 ? _a : {
                target: "word/document.xml",
                type: relationship_1.RelationshipTypes.OfficeDocument
            }, target = _b.target, type = _b.type;
            return d.loadRelationshipPart(target, type).then(function () { return d; });
        });
    };
    WordDocument.prototype.loadRelationshipPart = function (path, type) {
        var _this = this;
        if (this.partsMap[path])
            return Promise.resolve(this.partsMap[path]);
        if (!this._package.exists(path))
            return Promise.resolve(null);
        var part = null;
        switch (type) {
            case relationship_1.RelationshipTypes.OfficeDocument:
                this.documentPart = part = new document_part_1.DocumentPart(path, this._parser);
                break;
            case relationship_1.RelationshipTypes.FontTable:
                this.fontTablePart = part = new font_table_1.FontTablePart(path);
                break;
            case relationship_1.RelationshipTypes.Numbering:
                this.numberingPart = part = new numbering_part_1.NumberingPart(path, this._parser);
                break;
            case relationship_1.RelationshipTypes.Styles:
                this.stylesPart = part = new styles_part_1.StylesPart(path, this._parser);
                break;
        }
        if (part == null)
            return Promise.resolve(null);
        this.partsMap[path] = part;
        this.parts.push(part);
        return part.load(this._package).then(function () {
            if (part.rels == null || part.rels.length == 0)
                return part;
            var folder = utils_1.splitPath(part.path)[0];
            var rels = part.rels.map(function (rel) {
                return _this.loadRelationshipPart("" + folder + rel.target, rel.type);
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
    WordDocument.prototype.loadResource = function (part, id, outputType) {
        var rel = part.rels.find(function (x) { return x.id == id; });
        if (rel == null)
            return Promise.resolve(null);
        var fodler = utils_1.splitPath(part.path)[0];
        return this._package.load(fodler + rel.target, outputType);
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
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_jszip__;

/***/ })

/******/ });
});
//# sourceMappingURL=docx-preview.js.map