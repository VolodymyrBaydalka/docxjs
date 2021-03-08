import { DocxElement, IDomNumbering, NumberingPicBullet } from './document/dom';
import * as utils from './utils';
import { DocumentElement } from './document/document';
import { ParagraphElement, parseParagraphProperties, parseParagraphProperty } from './document/paragraph';
import { parseSectionProperties } from './document/section';
import globalXmlParser from './parser/xml-parser';
import { parseRunProperties, RunElement } from './document/run';
import { IDomStyle, IDomSubStyle } from './document/style';
import { BodyElement } from './document/body';
import { BreakElement } from './document/break';
import { HyperlinkElement } from './document/hyperlink';
import { TableCellElement } from './document/table-cell';
import { TableColumn, TableElement } from './document/table';
import { DrawingElement } from './document/drawing';
import { TableRowElement } from './document/table-row';
import { ImageElement } from './document/image';
import { deserializeElement } from './parser/xml-serialize';
import { FooterElement } from './footer/footer';
import { HeaderElement } from './header/header';

export var autos = {
    shd: "white",
    color: "black",
    highlight: "transparent"
};

export class DocumentParser {
    // removes XML declaration 
    skipDeclaration: boolean = true;

    // ignores page and table sizes
    ignoreWidth: boolean = false;
    debug: boolean = false;
    keepOrigin: boolean = false;

    private deserialize<T>(elem: Element, output: T): T {
        return deserializeElement(elem, output, { keepOrigin: this.keepOrigin });
    }

    parseDocumentFile(xmlDoc: Element): DocumentElement {
        const result = new DocumentElement();

        result.body = new BodyElement();

        var xbody = globalXmlParser.element(xmlDoc, "body");

        xml.foreach(xbody, elem => {
            switch (elem.localName) {
                case "p":
                    result.body.children.push(this.parseParagraph(elem));
                    break;

                case "tbl":
                    result.body.children.push(this.parseTable(elem));
                    break;

                case "sectPr":
                    result.body.sectionProps = parseSectionProperties(elem, globalXmlParser);
                    break;
            }
        });

        return result;
    }

    parseFooter(xmlDoc: Element): FooterElement {
        const result = new FooterElement();
    
        xml.foreach(xmlDoc, elem => {
            switch (elem.localName) {
                case "p":
                    result.children.push(this.parseParagraph(elem));
                    break;

                case "tbl":
                    result.children.push(this.parseTable(elem));
                    break;
            }
        });

        return result;
    }

    parseHeader(xmlDoc: Element): HeaderElement {
        const result = new HeaderElement();
    
        xml.foreach(xmlDoc, elem => {
            switch (elem.localName) {
                case "p":
                    result.children.push(this.parseParagraph(elem));
                    break;

                case "tbl":
                    result.children.push(this.parseTable(elem));
                    break;
            }
        });

        return result;
    }

    parseStylesFile(xstyles: Element): IDomStyle[] {
        var result = [];

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

    parseDefaultStyles(node: Element): IDomStyle {
        var result = <IDomStyle>{
            id: null,
            name: null,
            target: null,
            basedOn: null,
            styles: []
        };

        xml.foreach(node, c => {
            switch (c.localName) {
                case "rPrDefault":
                    var rPr = globalXmlParser.element(c, "rPr");

                    if (rPr)
                        result.styles.push({
                            target: "span",
                            values: this.parseDefaultProperties(rPr, {})
                        });
                    break;

                case "pPrDefault":
                    var pPr = globalXmlParser.element(c, "pPr");

                    if (pPr)
                        result.styles.push({
                            target: "p",
                            values: this.parseDefaultProperties(pPr, {})
                        });
                    break;
            }
        });

        return result;
    }

    parseStyle(node: Element): IDomStyle {
        var result = <IDomStyle>{
            id: xml.stringAttr(node, "styleId"),
            isDefault: xml.boolAttr(node, "default"),
            name: null,
            target: null,
            basedOn: null,
            styles: [],
            linked: null
        };

        switch (xml.stringAttr(node, "type")) {
            case "paragraph": result.target = "p"; break;
            case "table": result.target = "table"; break;
            case "character": result.target = "span"; break;
        }

        xml.foreach(node, n => {
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
                        values: this.parseDefaultProperties(n, {})
                    });
                    result.paragraphProps = parseParagraphProperties(n, globalXmlParser);
                    break;

                case "rPr":
                    result.styles.push({
                        target: "span",
                        values: this.parseDefaultProperties(n, {})
                    });
                    result.runProps = parseRunProperties(n, globalXmlParser);
                    break;

                case "tblPr":
                case "tcPr":
                    result.styles.push({
                        target: "td", //TODO: maybe move to processor
                        values: this.parseDefaultProperties(n, {})
                    });
                    break;

                case "tblStylePr":
                    for (let s of this.parseTableStyle(n))
                        result.styles.push(s);
                    break;

                case "rsid":
                case "qFormat":
                case "hidden":
                case "semiHidden":
                case "unhideWhenUsed":
                case "autoRedefine":
                case "uiPriority":
                    //TODO: ignore
                    break;

                default:
                    this.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
            }
        });

        return result;
    }

    parseTableStyle(node: Element): IDomSubStyle[] {
        var result = [];

        var type = xml.stringAttr(node, "type");
        var selector = "";

        switch (type) {
            case "firstRow": selector = "tr.first-row td"; break;
            case "lastRow": selector = "tr.last-row td"; break;
            case "firstCol": selector = "td.first-col"; break;
            case "lastCol": selector = "td.last-col"; break;
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

    parseNumberingFile(xnums: Element): IDomNumbering[] {
        var result = [];
        var mapping = {};
        var bullets = [];

        xml.foreach(xnums, n => {
            switch (n.localName) {
                case "abstractNum":
                    this.parseAbstractNumbering(n, bullets)
                        .forEach(x => result.push(x));
                    break;

                case "numPicBullet":
                    bullets.push(this.parseNumberingPicBullet(n));
                    break;

                case "num":
                    var numId = xml.stringAttr(n, "numId");
                    var abstractNumId = xml.elementStringAttr(n, "abstractNumId", "val");
                    mapping[abstractNumId] = numId;
                    break;
            }
        });

        result.forEach(x => x.id = mapping[x.id]);

        return result;
    }

    parseNumberingPicBullet(elem: Element): NumberingPicBullet {
        var pict = globalXmlParser.element(elem, "pict");
        var shape = pict && globalXmlParser.element(pict, "shape");
        var imagedata = shape && globalXmlParser.element(shape, "imagedata");

        return imagedata ? {
            id: xml.intAttr(elem, "numPicBulletId"),
            src: xml.stringAttr(imagedata, "id"),
            style: xml.stringAttr(shape, "style")
        } : null;
    }

    parseAbstractNumbering(node: Element, bullets: any[]): IDomNumbering[] {
        var result = [];
        var id = xml.stringAttr(node, "abstractNumId");

        xml.foreach(node, n => {
            switch (n.localName) {
                case "lvl":
                    result.push(this.parseNumberingLevel(id, n, bullets));
                    break;
            }
        });

        return result;
    }

    parseNumberingLevel(id: string, node: Element, bullets: any[]): IDomNumbering {
        var result: IDomNumbering = {
            id: id,
            level: xml.intAttr(node, "ilvl"),
            style: {}
        };

        xml.foreach(node, n => {
            switch (n.localName) {
                case "pPr":
                    this.parseDefaultProperties(n, result.style);
                    break;

                case "lvlPicBulletId":
                    var id = xml.intAttr(n, "val");
                    result.bullet = bullets.filter(x => x.id == id)[0];
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
    }


    parseParagraph(node: Element): ParagraphElement {
        const result = this.deserialize(node, new ParagraphElement());

        xml.foreach(node, c => {
            switch (c.localName) {
                case "r":
                    result.children.push(this.parseRun(c, result));
                    break;

                case "hyperlink":
                    result.children.push(this.parseHyperlink(c, result));
                    break;

                case "pPr":
                    this.parseParagraphProperties(c, result);
                    break;
            }
        });

        return result;
    }

    parseParagraphProperties(elem: Element, paragraph: ParagraphElement) {
        this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
            if(parseParagraphProperty(c, paragraph.props, globalXmlParser))
                return true;

            switch (c.localName) {
                case "pStyle":
                    utils.addElementClass(paragraph, xml.className(c, "val"));
                    break;

                case "cnfStyle":
                    utils.addElementClass(paragraph, values.classNameOfCnfStyle(c));
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

    parseFrame(node: Element, paragraph: ParagraphElement) {
        var dropCap = xml.stringAttr(node, "dropCap");

        if (dropCap == "drop")
            paragraph.cssStyle["float"] = "left";
    }

    parseHyperlink(node: Element, parent?: DocxElement): HyperlinkElement {
        var result = this.deserialize(node, new HyperlinkElement(parent));

        xml.foreach(node, c => {
            switch (c.localName) {
                case "r":
                    result.children.push(this.parseRun(c, result));
                    break;
            }
        });

        return result;
    }

    parseRun(node: Element, parent?: DocxElement): RunElement {
        var result = this.deserialize(node, new RunElement(parent));

        xml.foreach(node, c => {
            switch (c.localName) {
                case "lastRenderedPageBreak": {
                    const breakElem = new BreakElement();
                    breakElem.type = 'page';
                    result.children.push(breakElem);
                }
                    break;

                case "drawing":
                    let d = this.parseDrawing(c);

                    if (d)
                        result.children = [d];
                    break;

                case "rPr":
                    this.parseRunProperties(c, result);
                    break;
            }
        });

        return result;
    }

    parseRunProperties(elem: Element, run: RunElement) {

        Object.assign(run.props, parseRunProperties(elem, globalXmlParser));

        this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
            switch (c.localName) {
                case "rStyle":
                    run.className = xml.className(c, "val");
                    break;

                default:
                    return false;
            }

            return true;
        });
    }

    parseDrawing(node: Element): DocxElement {
        for (var n of globalXmlParser.elements(node)) {
            switch (n.localName) {
                case "inline":
                case "anchor":
                    return this.parseDrawingWrapper(n);
            }
        }
    }

    parseDrawingWrapper(node: Element): DocxElement {
        var result = new DrawingElement();
        var isAnchor = node.localName == "anchor";

        //TODO
        // result.style["margin-left"] = xml.sizeAttr(node, "distL", SizeType.Emu);
        // result.style["margin-top"] = xml.sizeAttr(node, "distT", SizeType.Emu);
        // result.style["margin-right"] = xml.sizeAttr(node, "distR", SizeType.Emu);
        // result.style["margin-bottom"] = xml.sizeAttr(node, "distB", SizeType.Emu);

        let wrapType: "wrapTopAndBottom" | "wrapNone" | null = null; 
        let simplePos = xml.boolAttr(node, "simplePos");

        let posX = { relative: "page", align: "left", offset: "0" };
        let posY = { relative: "page", align: "top", offset: "0" };

        for (var n of globalXmlParser.elements(node)) {
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
                        let pos = n.localName == "positionH" ? posX : posY;
                        var alignNode = globalXmlParser.element(n, "align");
                        var offsetNode = globalXmlParser.element(n, "posOffset");

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
        else if(wrapType == "wrapNone") {
            result.cssStyle['display'] = 'block';
            result.cssStyle['position'] = 'relative';
            result.cssStyle["width"] = "0px";
            result.cssStyle["height"] = "0px";

            if(posX.offset)
                result.cssStyle["left"] = posX.offset;
            if(posY.offset)
                result.cssStyle["top"] = posY.offset;
        }
        else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
            result.cssStyle["float"] = posX.align;
        }

        return result;
    }

    parseGraphic(elem: Element): DocxElement {
        var graphicData = globalXmlParser.element(elem, "graphicData");

        for (let n of globalXmlParser.elements(graphicData)) {
            switch (n.localName) {
                case "pic":
                    return this.parsePicture(n);
            }
        }

        return null;
    }

    parsePicture(elem: Element): ImageElement {
        var result = new ImageElement();
        var blipFill = globalXmlParser.element(elem, "blipFill");
        var blip = globalXmlParser.element(blipFill, "blip");

        result.src = xml.stringAttr(blip, "embed");

        var spPr = globalXmlParser.element(elem, "spPr");
        var xfrm = globalXmlParser.element(spPr, "xfrm");

        result.cssStyle["position"] = "relative";

        for (var n of globalXmlParser.elements(xfrm)) {
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
    }

    parseTable(node: Element): TableElement {
        var result = this.deserialize(node, new TableElement());

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

    parseTableColumns(node: Element): TableColumn[] {
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

    parseTableProperties(elem: Element, table: TableElement) {
        table.cssStyle = {};
        table.cellStyle = {};

        this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
            switch (c.localName) {
                case "tblStyle":
                    table.className = xml.className(c, "val");
                    break;

                case "tblLook":
                    utils.addElementClass(table, values.classNameOftblLook(c));
                    break;

                case "tblpPr":
                    this.parseTablePosition(c, table);
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
    }

    parseTablePosition(node: Element, table: TableElement) {
        var topFromText = xml.sizeAttr(node, "topFromText");
        var bottomFromText = xml.sizeAttr(node, "bottomFromText");
        var rightFromText = xml.sizeAttr(node, "rightFromText");
        var leftFromText = xml.sizeAttr(node, "leftFromText");

        table.cssStyle["float"] = 'left';
        table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
        table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
        table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
        table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
    }

    parseTableRow(node: Element): TableRowElement {
        var result = this.deserialize(node, new TableRowElement());

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

    parseTableRowProperties(elem: Element, row: TableRowElement) {
        row.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
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

    parseTableCell(node: Element): TableCellElement {
        var result = this.deserialize(node, new TableCellElement());

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

    parseTableCellProperties(elem: Element, cell: TableCellElement) {
        cell.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
            switch (c.localName) {
                case "gridSpan":
                    cell.span = xml.intAttr(c, "val", null);
                    break;

                case "vMerge": //TODO
                    cell.verticalMerge = xml.sizeAttr(c, "val");
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

    parseDefaultProperties(elem: Element, style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (prop: Element) => boolean = null): Record<string, string> {
        style = style || {};


        xml.foreach(elem, c => {
            switch (c.localName) {
                case "jc":
                    style["text-align"] = values.valueOfJc(c);
                    break;

                case "textAlignment":
                    style["vertical-align"] = values.valueOfTextAlignment(c);
                    break;

                case "color":
                    style["color"] = xml.colorAttr(c, "val", null, autos.color);
                    break;

                case "sz":
                    style["font-size"] = style["min-height"] = xml.sizeAttr(c, "val", SizeType.FontSize);
                    break;

                case "shd":
                    style["background-color"] = xml.colorAttr(c, "fill", null, autos.shd);
                    break;

                case "highlight":
                    style["background-color"] = xml.colorAttr(c, "val", null, autos.highlight);
                    break;

                case "tcW":
                    if (this.ignoreWidth)
                        break;

                case "tblW":
                    style["width"] = values.valueOfSize(c, "w");
                    break;

                case "trHeight":
                    this.parseTrHeight(c, style);
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
                    this.parseUnderline(c, style);
                    break;

                case "ind":
                case "tblInd":
                    this.parseIndentation(c, style);
                    break;

                case "rFonts":
                    this.parseFont(c, style);
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
                
                case "bdr":
                    style["border"] = values.valueOfBorder(c);
                    break;

                case "tcBorders":
                    this.parseBorderProperties(c, style);
                    break;

                case "noWrap":
                    //TODO
                    //style["white-space"] = "nowrap";
                    break;

                case "tblCellMar":
                case "tcMar":
                    this.parseMarginProperties(c, childStyle || style);
                    break;

                case "tblLayout":
                    style["table-layout"] = values.valueOfTblLayout(c);
                    break;

                case "vAlign":
                    style["vertical-align"] = xml.stringAttr(c, "val");
                    break;

                case "spacing":
                    if (elem.localName == "pPr")
                        this.parseSpacing(c, style);
                    break;

                case "lang":
                case "noProof": //ignore spellcheck
                case "webHidden": // maybe web-hidden should be implemented
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

    parseUnderline(node: Element, style: Record<string, string>) {
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
    }

    parseFont(node: Element, style: Record<string, string>) {
        var ascii = xml.stringAttr(node, "ascii");

        if (ascii)
            style["font-family"] = ascii;
    }

    parseIndentation(node: Element, style: Record<string, string>) {
        var firstLine = xml.sizeAttr(node, "firstLine");
        var left = xml.sizeAttr(node, "left");
        var start = xml.sizeAttr(node, "start");
        var right = xml.sizeAttr(node, "right");
        var end = xml.sizeAttr(node, "end");

        if (firstLine) style["text-indent"] = firstLine;
        if (left || start) style["margin-left"] = left || start;
        if (right || end) style["margin-right"] = right || end;
    }

    parseSpacing(node: Element, style: Record<string, string>) {
        var before = xml.sizeAttr(node, "before");
        var after = xml.sizeAttr(node, "after");
        var line = xml.intAttr(node, "line", null);
        var lineRule = xml.stringAttr(node, "lineRule");

        if (before) style["margin-top"] = before;
        if (after) style["margin-bottom"] = after;
        
        if (line !== null) {
            switch(lineRule) {
                case "auto": 
                    style["line-height"] = `${(line / 240).toFixed(2)}`;
                    break;

                case "atLeast":
                    style["line-height"] = `calc(100% + ${line / 20}pt)`;
                    break;

                default:
                    style["line-height"] = style["min-height"] = `${line / 20}pt`
                    break;
            }
        }
    }

    parseMarginProperties(node: Element, output: Record<string, string>) {
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

    parseTrHeight(node: Element, output: Record<string, string>) {
        switch (xml.stringAttr(node, "hRule")) {
            case "exact":
                output["height"] = xml.sizeAttr(node, "val");
                break;

            case "atLeast":
            default:
                output["height"] = xml.sizeAttr(node, "val");
                // min-height doesn't work for tr
                //output["min-height"] = xml.sizeAttr(node, "val");  
                break;
        }
    }

    parseBorderProperties(node: Element, output: Record<string, string>) {
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
    Emu,
    Border,
    Percent
}

class xml {
    static foreach(node: Element, cb: (n: Element) => void) {
        for (var i = 0; i < node.childNodes.length; i++) {
            let n = node.childNodes[i];

            if (n.nodeType == 1)
                cb(<Element>n);
        }
    }

    static elementStringAttr(elem: Element, nodeName, attrName: string) {
        var n = globalXmlParser.element(elem, nodeName)
        return n ? xml.stringAttr(n, attrName) : null;
    }

    static stringAttr(node: Element, attrName: string) {
        return globalXmlParser.attr(node, attrName);
    }

    static colorAttr(node: Element, attrName: string, defValue: string = null, autoColor: string = 'black') {
        var v = xml.stringAttr(node, attrName);

        switch (v) {
            case "yellow":
                return v;

            case "auto":
                return autoColor;
        }

        return v ? `#${v}` : defValue;
    }

    static boolAttr(node: Element, attrName: string, defValue: boolean = false) {
        return globalXmlParser.boolAttr(node, attrName, defValue);
    }

    static intAttr(node: Element, attrName: string, defValue: number = 0) {
        var val = xml.stringAttr(node, attrName);
        return val ? parseInt(xml.stringAttr(node, attrName)) : defValue;
    }

    static sizeAttr(node: Element, attrName: string, type: SizeType = SizeType.Dxa) {
        return xml.convertSize(xml.stringAttr(node, attrName), type);
    }

    static sizeValue(node: Element, type: SizeType = SizeType.Dxa) {
        return xml.convertSize(node.textContent, type);
    }

    static convertSize(val: string, type: SizeType = SizeType.Dxa) {
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
    }

    static className(node: Element, attrName: string) {
        var val = xml.stringAttr(node, attrName);

        return val && val.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and');
    }
}

class values {
    static valueOfBold(c: Element) {
        return xml.boolAttr(c, "val", true) ? "bold" : "normal"
    }

    static valueOfSize(c: Element, attr: string) {
        var type: SizeType = SizeType.Dxa;

        switch (xml.stringAttr(c, "type")) {
            case "dxa": break;
            case "pct": type = SizeType.Percent; break;
        }

        return xml.sizeAttr(c, attr, type);
    }

    static valueOfStrike(c: Element) {
        return xml.boolAttr(c, "val", true) ? "line-through" : "none"
    }

    static valueOfMargin(c: Element) {
        return xml.sizeAttr(c, "w");
    }

    static valueOfBorder(c: Element) {
        var type = xml.stringAttr(c, "val");

        if (type == "nil")
            return "none";

        var color = xml.colorAttr(c, "color");
        var size = xml.sizeAttr(c, "sz", SizeType.Border);

        return `${size} solid ${color == "auto" ? "black" : color}`;
    }

    static valueOfTblLayout(c: Element) {
        var type = xml.stringAttr(c, "val");
        return type == "fixed" ? "fixed" : "auto";
    }

    static classNameOfCnfStyle(c: Element) {
        let className = "";
        let val = xml.stringAttr(c, "val");
        //FirstRow, LastRow, FirstColumn, LastColumn, Band1Vertical, Band2Vertical, Band1Horizontal, Band2Horizontal, NE Cell, NW Cell, SE Cell, SW Cell.

        if (val[0] == "1") className += " first-row";
        if (val[1] == "1") className += " last-row";
        if (val[2] == "1") className += " first-col";
        if (val[3] == "1") className += " last-col";
        if (val[4] == "1") className += " odd-col";
        if (val[5] == "1") className += " even-col";
        if (val[6] == "1") className += " odd-row";
        if (val[7] == "1") className += " even-row";
        if (val[8] == "1") className += " ne-cell";
        if (val[9] == "1") className += " nw-cell";
        if (val[10] == "1") className += " se-cell";
        if (val[11] == "1") className += " sw-cell";

        return className.trim();
    }

    static valueOfJc(c: Element) {
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
    }

    static valueOfTextAlignment(c: Element) {
        var type = xml.stringAttr(c, "val");

        switch (type) {
            case "auto":
            case "baseline": return "baseline";
            case "top": return "top";
            case "center": return "middle";
            case "bottom": return "bottom";
        }

        return type;
    }

    static addSize(a: string, b: string): string {
        if (a == null) return b;
        if (b == null) return a;

        return `calc(${a} + ${b})`; //TODO
    }

    static classNameOftblLook(c: Element) {
        let className = "";

        if (xml.boolAttr(c, "firstColumn")) className += " first-col";
        if (xml.boolAttr(c, "firstRow")) className += " first-row";
        if (xml.boolAttr(c, "lastColumn")) className += " lat-col";
        if (xml.boolAttr(c, "lastRow")) className += " last-row";
        if (xml.boolAttr(c, "noHBand")) className += " no-hband";
        if (xml.boolAttr(c, "noVBand")) className += " no-vband";

        return className.trim();
    }
}