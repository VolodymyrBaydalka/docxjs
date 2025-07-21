import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, WmlSmartTag, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell,
	WmlTableRow, NumberingPicBullet, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlAltChunk
} from './document/dom';
import { DocumentElement } from './document/document';
import { WmlParagraph, parseParagraphProperties, parseParagraphProperty } from './document/paragraph';
import { parseSectionProperties, SectionProperties } from './document/section';
import xml from './parser/xml-parser';
import { parseRunProperties, WmlRun } from './document/run';
import { parseBookmarkEnd, parseBookmarkStart } from './document/bookmarks';
import { IDomStyle, IDomSubStyle } from './document/style';
import { WmlFieldChar, WmlFieldSimple, WmlInstructionText } from './document/fields';
import { convertLength, LengthUsage, LengthUsageType } from './document/common';
import { parseVmlElement } from './vml/vml';
import { WmlComment, WmlCommentRangeEnd, WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import { encloseFontFamily } from './utils';

export var autos = {
	shd: "inherit",
	color: "black",
	borderColor: "black",
	highlight: "transparent"
};

const supportedNamespaceURIs = [];

const mmlTagMap = {
	"oMath": DomType.MmlMath,
	"oMathPara": DomType.MmlMathParagraph,
	"f": DomType.MmlFraction,
	"func": DomType.MmlFunction,
	"fName": DomType.MmlFunctionName,
	"num": DomType.MmlNumerator,
	"den": DomType.MmlDenominator,
	"rad": DomType.MmlRadical,
	"deg": DomType.MmlDegree,
	"e": DomType.MmlBase,
	"sSup": DomType.MmlSuperscript,
	"sSub": DomType.MmlSubscript,
	"sPre": DomType.MmlPreSubSuper,
	"sup": DomType.MmlSuperArgument,
	"sub": DomType.MmlSubArgument,
	"d": DomType.MmlDelimiter,
	"nary": DomType.MmlNary,
	"eqArr": DomType.MmlEquationArray,
	"lim": DomType.MmlLimit,
	"limLow": DomType.MmlLimitLower,
	"m": DomType.MmlMatrix,
	"mr": DomType.MmlMatrixRow,
	"box": DomType.MmlBox,
	"bar": DomType.MmlBar,
	"groupChr": DomType.MmlGroupChar
}

export interface DocumentParserOptions {
	ignoreWidth: boolean;
	debug: boolean;
}

export class DocumentParser {
	options: DocumentParserOptions;

	constructor(options?: Partial<DocumentParserOptions>) {
		this.options = {
			ignoreWidth: false,
			debug: false,
			...options
		};
	}

	parseNotes(xmlDoc: Element, elemName: string, elemClass: any): any[] {
		var result = [];

		for (let el of xml.elements(xmlDoc, elemName)) {
			const node = new elemClass();
			node.id = xml.attr(el, "id");
			node.noteType = xml.attr(el, "type");
			node.children = this.parseBodyElements(el);
			result.push(node);
		}

		return result;
	}

	parseComments(xmlDoc: Element): any[] {
		var result = [];

		for (let el of xml.elements(xmlDoc, "comment")) {
			const item = new WmlComment();
			item.id = xml.attr(el, "id");
			item.author = xml.attr(el, "author");
			item.initials = xml.attr(el, "initials");
			item.date = xml.attr(el, "date");
			item.children = this.parseBodyElements(el);
			result.push(item);
		}

		return result;
	}

	parseDocumentFile(xmlDoc: Element): DocumentElement {
		var xbody = xml.element(xmlDoc, "body");
		var background = xml.element(xmlDoc, "background");
		var sectPr = xml.element(xbody, "sectPr");

		return {
			type: DomType.Document,
			children: this.parseBodyElements(xbody),
			props: sectPr ? parseSectionProperties(sectPr, xml) : {} as SectionProperties,
			cssStyle: background ? this.parseBackground(background) : {},
		};
	}

	parseBackground(elem: Element): any {
		var result = {};
		var color = xmlUtil.colorAttr(elem, "color");

		if (color) {
			result["background-color"] = color;
		}

		return result;
	}

	parseBodyElements(element: Element): OpenXmlElement[] {
		var children = [];

		for (let elem of xml.elements(element)) {
			switch (elem.localName) {
				case "p":
					children.push(this.parseParagraph(elem));
					break;

				case "altChunk":
					children.push(this.parseAltChunk(elem));
					break;
	
				case "tbl":
					children.push(this.parseTable(elem));
					break;

				case "sdt":
					children.push(...this.parseSdt(elem, e => this.parseBodyElements(e)));
					break;
			}
		}

		return children;
	}

	parseStylesFile(xstyles: Element): IDomStyle[] {
		var result = [];

		xmlUtil.foreach(xstyles, n => {
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

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "rPrDefault":
					var rPr = xml.element(c, "rPr");

					if (rPr)
						result.styles.push({
							target: "span",
							values: this.parseDefaultProperties(rPr, {})
						});
					break;

				case "pPrDefault":
					var pPr = xml.element(c, "pPr");

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
			id: xml.attr(node, "styleId"),
			isDefault: xml.boolAttr(node, "default"),
			name: null,
			target: null,
			basedOn: null,
			styles: [],
			linked: null
		};

		switch (xml.attr(node, "type")) {
			case "paragraph": result.target = "p"; break;
			case "table": result.target = "table"; break;
			case "character": result.target = "span"; break;
			//case "numbering": result.target = "p"; break;
		}

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "basedOn":
					result.basedOn = xml.attr(n, "val");
					break;

				case "name":
					result.name = xml.attr(n, "val");
					break;

				case "link":
					result.linked = xml.attr(n, "val");
					break;

				case "next":
					result.next = xml.attr(n, "val");
					break;

				case "aliases":
					result.aliases = xml.attr(n, "val").split(",");
					break;

				case "pPr":
					result.styles.push({
						target: "p",
						values: this.parseDefaultProperties(n, {})
					});
					result.paragraphProps = parseParagraphProperties(n, xml);
					break;

				case "rPr":
					result.styles.push({
						target: "span",
						values: this.parseDefaultProperties(n, {})
					});
					result.runProps = parseRunProperties(n, xml);
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
					this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
			}
		});

		return result;
	}

	parseTableStyle(node: Element): IDomSubStyle[] {
		var result = [];

		var type = xml.attr(node, "type");
		var selector = "";
		var modificator = "";

		switch (type) {
			case "firstRow":
				modificator = ".first-row";
				selector = "tr.first-row td";
				break;
			case "lastRow":
				modificator = ".last-row";
				selector = "tr.last-row td";
				break;
			case "firstCol":
				modificator = ".first-col";
				selector = "td.first-col";
				break;
			case "lastCol":
				modificator = ".last-col";
				selector = "td.last-col";
				break;
			case "band1Vert":
				modificator = ":not(.no-vband)";
				selector = "td.odd-col";
				break;
			case "band2Vert":
				modificator = ":not(.no-vband)";
				selector = "td.even-col";
				break;
			case "band1Horz":
				modificator = ":not(.no-hband)";
				selector = "tr.odd-row";
				break;
			case "band2Horz":
				modificator = ":not(.no-hband)";
				selector = "tr.even-row";
				break;
			default: return [];
		}

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "pPr":
					result.push({
						target: `${selector} p`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "rPr":
					result.push({
						target: `${selector} span`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "tblPr":
				case "tcPr":
					result.push({
						target: selector, //TODO: maybe move to processor
						mod: modificator,
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

		xmlUtil.foreach(xnums, n => {
			switch (n.localName) {
				case "abstractNum":
					this.parseAbstractNumbering(n, bullets)
						.forEach(x => result.push(x));
					break;

				case "numPicBullet":
					bullets.push(this.parseNumberingPicBullet(n));
					break;

				case "num":
					var numId = xml.attr(n, "numId");
					var abstractNumId = xml.elementAttr(n, "abstractNumId", "val");
					mapping[abstractNumId] = numId;
					break;
			}
		});

		result.forEach(x => x.id = mapping[x.id]);

		return result;
	}

	parseNumberingPicBullet(elem: Element): NumberingPicBullet {
		var pict = xml.element(elem, "pict");
		var shape = pict && xml.element(pict, "shape");
		var imagedata = shape && xml.element(shape, "imagedata");

		return imagedata ? {
			id: xml.intAttr(elem, "numPicBulletId"),
			src: xml.attr(imagedata, "id"),
			style: xml.attr(shape, "style")
		} : null;
	}

	parseAbstractNumbering(node: Element, bullets: any[]): IDomNumbering[] {
		var result = [];
		var id = xml.attr(node, "abstractNumId");

		xmlUtil.foreach(node, n => {
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
			start: 1,
			pStyleName: undefined,
			pStyle: {},
			rStyle: {},
			suff: "tab"
		};

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "start":
					result.start = xml.intAttr(n, "val");
					break;

				case "pPr":
					this.parseDefaultProperties(n, result.pStyle);
					break;

				case "rPr":
					this.parseDefaultProperties(n, result.rStyle);
					break;

				case "lvlPicBulletId":
					var id = xml.intAttr(n, "val");
					result.bullet = bullets.find(x => x?.id == id);
					break;

				case "lvlText":
					result.levelText = xml.attr(n, "val");
					break;

				case "pStyle":
					result.pStyleName = xml.attr(n, "val");
					break;

				case "numFmt":
					result.format = xml.attr(n, "val");
					break;

				case "suff":
					result.suff = xml.attr(n, "val");
					break;
			}
		});

		return result;
	}

	parseSdt(node: Element, parser: Function): OpenXmlElement[] {
		const sdtContent = xml.element(node, "sdtContent");
		return sdtContent ? parser(sdtContent) : [];
	}

	parseInserted(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{ 
			type: DomType.Inserted, 
			children: parentParser(node)?.children ?? []
		};
	}

	parseDeleted(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{ 
			type: DomType.Deleted, 
			children: parentParser(node)?.children ?? []
		};
	}

	parseAltChunk(node: Element): WmlAltChunk {
		return { type: DomType.AltChunk, children: [], id: xml.attr(node, "id") };
	}

	parseParagraph(node: Element): OpenXmlElement {
		var result = <WmlParagraph>{ type: DomType.Paragraph, children: [] };

		for (let el of xml.elements(node)) {
			switch (el.localName) {
				case "pPr":
					this.parseParagraphProperties(el, result);
					break;

				case "r":
					result.children.push(this.parseRun(el, result));
					break;

				case "hyperlink":
					result.children.push(this.parseHyperlink(el, result));
					break;
				
				case "smartTag":
					result.children.push(this.parseSmartTag(el, result));
					break;

				case "bookmarkStart":
					result.children.push(parseBookmarkStart(el, xml));
					break;

				case "bookmarkEnd":
					result.children.push(parseBookmarkEnd(el, xml));
					break;

				case "commentRangeStart":
					result.children.push(new WmlCommentRangeStart(xml.attr(el, "id")));
					break;
	
				case "commentRangeEnd":
					result.children.push(new WmlCommentRangeEnd(xml.attr(el, "id")));
					break;

				case "oMath":
				case "oMathPara":
					result.children.push(this.parseMathElement(el));
					break;

				case "sdt":
					result.children.push(...this.parseSdt(el, e => this.parseParagraph(e).children));
					break;

				case "ins":
					result.children.push(this.parseInserted(el, e => this.parseParagraph(e)));
					break;

				case "del":
					result.children.push(this.parseDeleted(el, e => this.parseParagraph(e)));
					break;
			}
		}

		return result;
	}

	parseParagraphProperties(elem: Element, paragraph: WmlParagraph) {
		this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
			if (parseParagraphProperty(c, paragraph, xml))
				return true;

			switch (c.localName) {
				case "pStyle":
					paragraph.styleName = xml.attr(c, "val");
					break;

				case "cnfStyle":
					paragraph.className = values.classNameOfCnfStyle(c);
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

	parseFrame(node: Element, paragraph: WmlParagraph) {
		var dropCap = xml.attr(node, "dropCap");

		if (dropCap == "drop")
			paragraph.cssStyle["float"] = "left";
	}

	parseHyperlink(node: Element, parent?: OpenXmlElement): WmlHyperlink {
		var result: WmlHyperlink = <WmlHyperlink>{ type: DomType.Hyperlink, parent: parent, children: [] };
		
		result.anchor = xml.attr(node, "anchor");
		result.id = xml.attr(node, "id");

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "r":
					result.children.push(this.parseRun(c, result));
					break;
			}
		});

		return result;
	}
	
	parseSmartTag(node: Element, parent?: OpenXmlElement): WmlSmartTag {
		var result: WmlSmartTag = { type: DomType.SmartTag, parent, children: [] };
		var uri = xml.attr(node, "uri");
		var element = xml.attr(node, "element");

		if (uri)
			result.uri = uri;

		if (element)
			result.element = element;

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "r":
					result.children.push(this.parseRun(c, result));
					break;
			}
		});

		return result;
	}

	parseRun(node: Element, parent?: OpenXmlElement): WmlRun {
		var result: WmlRun = <WmlRun>{ type: DomType.Run, parent: parent, children: [] };

		xmlUtil.foreach(node, c => {
			c = this.checkAlternateContent(c);

			switch (c.localName) {
				case "t":
					result.children.push(<WmlText>{
						type: DomType.Text,
						text: c.textContent
					});//.replace(" ", "\u00A0"); // TODO
					break;

				case "delText":
					result.children.push(<WmlText>{
						type: DomType.DeletedText,
						text: c.textContent
					});
					break;

				case "commentReference":
					result.children.push(new WmlCommentReference(xml.attr(c, "id")));
					break;

				case "fldSimple":
					result.children.push(<WmlFieldSimple>{
						type: DomType.SimpleField,
						instruction: xml.attr(c, "instr"),
						lock: xml.boolAttr(c, "lock", false),
						dirty: xml.boolAttr(c, "dirty", false)
					});
					break;

				case "instrText":
					result.fieldRun = true;
					result.children.push(<WmlInstructionText>{
						type: DomType.Instruction,
						text: c.textContent
					});
					break;

				case "fldChar":
					result.fieldRun = true;
					result.children.push(<WmlFieldChar>{
						type: DomType.ComplexField,
						charType: xml.attr(c, "fldCharType"),
						lock: xml.boolAttr(c, "lock", false),
						dirty: xml.boolAttr(c, "dirty", false)
					});
					break;

				case "noBreakHyphen":
					result.children.push({ type: DomType.NoBreakHyphen });
					break;

				case "br":
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: xml.attr(c, "type") || "textWrapping"
					});
					break;

				case "lastRenderedPageBreak":
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: "lastRenderedPageBreak"
					});
					break;

				case "sym":
					result.children.push(<WmlSymbol>{
						type: DomType.Symbol,
						font: encloseFontFamily(xml.attr(c, "font")),
						char: xml.attr(c, "char")
					});
					break;

				case "tab":
					result.children.push({ type: DomType.Tab });
					break;

				case "footnoteReference":
					result.children.push(<WmlNoteReference>{
						type: DomType.FootnoteReference,
						id: xml.attr(c, "id")
					});
					break;

				case "endnoteReference":
					result.children.push(<WmlNoteReference>{
						type: DomType.EndnoteReference,
						id: xml.attr(c, "id")
					});
					break;

				case "drawing":
					let d = this.parseDrawing(c);

					if (d)
						result.children = [d];
					break;

				case "pict":
					result.children.push(this.parseVmlPicture(c));
					break;

				case "rPr":
					this.parseRunProperties(c, result);
					break;
			}
		});

		return result;
	}

	parseMathElement(elem: Element): OpenXmlElement {
		const propsTag = `${elem.localName}Pr`;
		const result = { type: mmlTagMap[elem.localName], children: [] } as OpenXmlElement;

		for (const el of xml.elements(elem)) {
			const childType = mmlTagMap[el.localName];

			if (childType) {
				result.children.push(this.parseMathElement(el));
			} else if (el.localName == "r") {
				var run = this.parseRun(el);
				run.type = DomType.MmlRun;
				result.children.push(run);
			} else if (el.localName == propsTag) {
				result.props = this.parseMathProperies(el);
			}
		}

		return result;
	}

	parseMathProperies(elem: Element): Record<string, any> {
		const result: Record<string, any> = {};

		for (const el of xml.elements(elem)) {
			switch (el.localName) {
				case "chr": result.char = xml.attr(el, "val"); break;
				case "vertJc": result.verticalJustification = xml.attr(el, "val"); break;
				case "pos": result.position = xml.attr(el, "val"); break;
				case "degHide": result.hideDegree = xml.boolAttr(el, "val"); break;
				case "begChr": result.beginChar = xml.attr(el, "val"); break;
				case "endChr": result.endChar = xml.attr(el, "val"); break;
			}
		}

		return result;
	}

	parseRunProperties(elem: Element, run: WmlRun) {
		this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
			switch (c.localName) {
				case "rStyle":
					run.styleName = xml.attr(c, "val");
					break;

				case "vertAlign":
					run.verticalAlign = values.valueOfVertAlign(c, true);
					break;

				default:
					return false;
			}

			return true;
		});
	}

	parseVmlPicture(elem: Element): OpenXmlElement {
		const result = { type: DomType.VmlPicture, children: [] };

		for (const el of xml.elements(elem)) {
			const child = parseVmlElement(el, this);
			child && result.children.push(child);
		}

		return result;
	}

	checkAlternateContent(elem: Element): Element {
		if (elem.localName != 'AlternateContent')
			return elem;

		var choice = xml.element(elem, "Choice");

		if (choice) {
			var requires = xml.attr(choice, "Requires");
			var namespaceURI = elem.lookupNamespaceURI(requires);

			if (supportedNamespaceURIs.includes(namespaceURI))
				return choice.firstElementChild;
		}

		return xml.element(elem, "Fallback")?.firstElementChild;
	}

	parseDrawing(node: Element): OpenXmlElement {
		for (var n of xml.elements(node)) {
			switch (n.localName) {
				case "inline":
				case "anchor":
					return this.parseDrawingWrapper(n);
			}
		}
	}

	parseDrawingWrapper(node: Element): OpenXmlElement {
		var result = <OpenXmlElement>{ type: DomType.Drawing, children: [], cssStyle: {} };
		var isAnchor = node.localName == "anchor";

		//TODO
		// result.style["margin-left"] = xml.sizeAttr(node, "distL", SizeType.Emu);
		// result.style["margin-top"] = xml.sizeAttr(node, "distT", SizeType.Emu);
		// result.style["margin-right"] = xml.sizeAttr(node, "distR", SizeType.Emu);
		// result.style["margin-bottom"] = xml.sizeAttr(node, "distB", SizeType.Emu);

		let wrapType: "wrapTopAndBottom" | "wrapNone" | null = null;
		let simplePos = xml.boolAttr(node, "simplePos");
		let behindDoc = xml.boolAttr(node, "behindDoc");

		let posX = { relative: "page", align: "left", offset: "0" };
		let posY = { relative: "page", align: "top", offset: "0" };

		for (var n of xml.elements(node)) {
			switch (n.localName) {
				case "simplePos":
					if (simplePos) {
						posX.offset = xml.lengthAttr(n, "x", LengthUsage.Emu);
						posY.offset = xml.lengthAttr(n, "y", LengthUsage.Emu);
					}
					break;

				case "extent":
					result.cssStyle["width"] = xml.lengthAttr(n, "cx", LengthUsage.Emu);
					result.cssStyle["height"] = xml.lengthAttr(n, "cy", LengthUsage.Emu);
					break;

				case "positionH":
				case "positionV":
					if (!simplePos) {
						let pos = n.localName == "positionH" ? posX : posY;
						var alignNode = xml.element(n, "align");
						var offsetNode = xml.element(n, "posOffset");

						pos.relative = xml.attr(n, "relativeFrom") ?? pos.relative;

						if (alignNode)
							pos.align = alignNode.textContent;

						if (offsetNode)
							pos.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu);
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
	}

	parseGraphic(elem: Element): OpenXmlElement {
		var graphicData = xml.element(elem, "graphicData");

		for (let n of xml.elements(graphicData)) {
			switch (n.localName) {
				case "pic":
					return this.parsePicture(n);
			}
		}

		return null;
	}

	parsePicture(elem: Element): IDomImage {
		var result = <IDomImage>{ type: DomType.Image, src: "", cssStyle: {} };
		var blipFill = xml.element(elem, "blipFill");
		var blip = xml.element(blipFill, "blip");

		result.src = xml.attr(blip, "embed");

		var spPr = xml.element(elem, "spPr");
		var xfrm = xml.element(spPr, "xfrm");

		result.cssStyle["position"] = "relative";

		for (var n of xml.elements(xfrm)) {
			switch (n.localName) {
				case "ext":
					result.cssStyle["width"] = xml.lengthAttr(n, "cx", LengthUsage.Emu);
					result.cssStyle["height"] = xml.lengthAttr(n, "cy", LengthUsage.Emu);
					break;

				case "off":
					result.cssStyle["left"] = xml.lengthAttr(n, "x", LengthUsage.Emu);
					result.cssStyle["top"] = xml.lengthAttr(n, "y", LengthUsage.Emu);
					break;
			}
		}

		return result;
	}

	parseTable(node: Element): WmlTable {
		var result: WmlTable = { type: DomType.Table, children: [] };

		xmlUtil.foreach(node, c => {
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

	parseTableColumns(node: Element): WmlTableColumn[] {
		var result = [];

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "gridCol":
					result.push({ width: xml.lengthAttr(n, "w") });
					break;
			}
		});

		return result;
	}

	parseTableProperties(elem: Element, table: WmlTable) {
		table.cssStyle = {};
		table.cellStyle = {};

		this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
			switch (c.localName) {
				case "tblStyle":
					table.styleName = xml.attr(c, "val");
					break;

				case "tblLook":
					table.className = values.classNameOftblLook(c);
					break;

				case "tblpPr":
					this.parseTablePosition(c, table);
					break;

				case "tblStyleColBandSize":
					table.colBandSize = xml.intAttr(c, "val");
					break;

				case "tblStyleRowBandSize":
					table.rowBandSize = xml.intAttr(c, "val");
					break;

					
				case "hidden":
					table.cssStyle["display"] = "none";
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

	parseTablePosition(node: Element, table: WmlTable) {
		var topFromText = xml.lengthAttr(node, "topFromText");
		var bottomFromText = xml.lengthAttr(node, "bottomFromText");
		var rightFromText = xml.lengthAttr(node, "rightFromText");
		var leftFromText = xml.lengthAttr(node, "leftFromText");

		table.cssStyle["float"] = 'left';
		table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
		table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
		table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
		table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
	}

	parseTableRow(node: Element): WmlTableRow {
		var result: WmlTableRow = { type: DomType.Row, children: [] };

		xmlUtil.foreach(node, c => {
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

	parseTableRowProperties(elem: Element, row: WmlTableRow) {
		row.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
			switch (c.localName) {
				case "cnfStyle":
					row.className = values.classNameOfCnfStyle(c);
					break;

				case "tblHeader":
					row.isHeader = xml.boolAttr(c, "val");
					break;

				case "gridBefore":
					row.gridBefore = xml.intAttr(c, "val");
					break;

				case "gridAfter":
					row.gridAfter = xml.intAttr(c, "val");
					break;

				default:
					return false;
			}

			return true;
		});
	}

	parseTableCell(node: Element): OpenXmlElement {
		var result: WmlTableCell = { type: DomType.Cell, children: [] };

		xmlUtil.foreach(node, c => {
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

	parseTableCellProperties(elem: Element, cell: WmlTableCell) {
		cell.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
			switch (c.localName) {
				case "gridSpan":
					cell.span = xml.intAttr(c, "val", null);
					break;

				case "vMerge":
					cell.verticalMerge = xml.attr(c, "val") ?? "continue";
					break;

				case "cnfStyle":
					cell.className = values.classNameOfCnfStyle(c);
					break;

				default:
					return false;
			}

			return true;
		});

		this.parseTableCellVerticalText(elem, cell);
	}

	parseTableCellVerticalText(elem: Element, cell: WmlTableCell) {
		const directionMap = {
			"btLr": {
				writingMode: "vertical-rl",
				transform: "rotate(180deg)"
			},
			"lrTb": {
				writingMode: "vertical-lr",
				transform: "none"
			},
			"tbRl": {
				writingMode: "vertical-rl",
				transform: "none"
			}
		};

		xmlUtil.foreach(elem, c => {
			if (c.localName === "textDirection") {
				const direction = xml.attr(c, "val");
				const style = directionMap[direction] || {writingMode: "horizontal-tb"};
				cell.cssStyle["writing-mode"] = style.writingMode;
				cell.cssStyle["transform"] = style.transform;
			}
		});
	}

	parseDefaultProperties(elem: Element, style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (prop: Element) => boolean = null): Record<string, string> {
		style = style || {};

		xmlUtil.foreach(elem, c => {
			if (handler?.(c))
				return;

			switch (c.localName) {
				case "jc":
					style["text-align"] = values.valueOfJc(c);
					break;

				case "textAlignment":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "color":
					style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
					break;

				case "sz":
					style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				case "shd":
					style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
					break;

				case "highlight":
					style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
					break;

				case "vertAlign":
					//TODO
					// style.verticalAlign = values.valueOfVertAlign(c);
					break;

				case "position":
					style.verticalAlign = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				case "tcW":
					if (this.options.ignoreWidth)
						break;

				case "tblW":
					style["width"] = values.valueOfSize(c, "w");
					break;

				case "trHeight":
					this.parseTrHeight(c, style);
					break;

				case "strike":
					style["text-decoration"] = xml.boolAttr(c, "val", true) ? "line-through" : "none"
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
					style["font-variant"] = xml.boolAttr(c, "val", true) ? "small-caps" : "none";
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

				case "vanish":
					if (xml.boolAttr(c, "val", true))
						style["display"] = "none";
					break;

				case "kern":
					//TODO
					//style['letter-spacing'] = xml.lengthAttr(elem, 'val', LengthUsage.FontSize);
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
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "spacing":
					if (elem.localName == "pPr")
						this.parseSpacing(c, style);
					break;

				case "wordWrap":
					if (xml.boolAttr(c, "val")) //TODO: test with examples
						style["overflow-wrap"] = "break-word";
					break;

				case "suppressAutoHyphens":
					style["hyphens"] = xml.boolAttr(c, "val", true) ? "none" : "auto";
					break;

				case "lang":
					style["$lang"] = xml.attr(c, "val");
					break;

				case "rtl":
				case "bidi":
					if (xml.boolAttr(c, "val", true))
						style["direction"] = "rtl";
					break;

				case "bCs":
				case "iCs":
				case "szCs":
				case "tabs": //ignore - tabs is parsed by other parser
				case "outlineLvl": //TODO
				case "contextualSpacing": //TODO
				case "tblStyleColBandSize": //TODO
				case "tblStyleRowBandSize": //TODO
				case "webHidden": //TODO - maybe web-hidden should be implemented
				case "pageBreakBefore": //TODO - maybe ignore 
				case "suppressLineNumbers": //TODO - maybe ignore
				case "keepLines": //TODO - maybe ignore
				case "keepNext": //TODO - maybe ignore
				case "widowControl": //TODO - maybe ignore 
				case "bidi": //TODO - maybe ignore
				case "rtl": //TODO - maybe ignore
				case "noProof": //ignore spellcheck
					//TODO ignore
					break;

				default:
					if (this.options.debug)
						console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
					break;
			}
		});

		return style;
	}

	parseUnderline(node: Element, style: Record<string, string>) {
		var val = xml.attr(node, "val");

		if (val == null)
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
				style["text-decoration"] = "underline dashed";
				break;

			case "dotted":
			case "dottedHeavy":
				style["text-decoration"] = "underline dotted";
				break;

			case "double":
				style["text-decoration"] = "underline double";
				break;

			case "single":
			case "thick":
				style["text-decoration"] = "underline";
				break;

			case "wave":
			case "wavyDouble":
			case "wavyHeavy":
				style["text-decoration"] = "underline wavy";
				break;

			case "words":
				style["text-decoration"] = "underline";
				break;

			case "none":
				style["text-decoration"] = "none";
				break;
		}

		var col = xmlUtil.colorAttr(node, "color");

		if (col)
			style["text-decoration-color"] = col;
	}

	parseFont(node: Element, style: Record<string, string>) {
		var ascii = xml.attr(node, "ascii");
		var asciiTheme = values.themeValue(node, "asciiTheme");
		var eastAsia = xml.attr(node, "eastAsia");
		var fonts = [ascii, asciiTheme, eastAsia].filter(x => x).map(x => encloseFontFamily(x));

		if (fonts.length > 0)
			style["font-family"] = [...new Set(fonts)].join(', ');
	}

	parseIndentation(node: Element, style: Record<string, string>) {
		var firstLine = xml.lengthAttr(node, "firstLine");
		var hanging = xml.lengthAttr(node, "hanging");
		var left = xml.lengthAttr(node, "left");
		var start = xml.lengthAttr(node, "start");
		var right = xml.lengthAttr(node, "right");
		var end = xml.lengthAttr(node, "end");

		if (firstLine) style["text-indent"] = firstLine;
		if (hanging) style["text-indent"] = `-${hanging}`;
		if (left || start) style["margin-inline-start"] = left || start;
		if (right || end) style["margin-inline-end"] = right || end;
	}

	parseSpacing(node: Element, style: Record<string, string>) {
		var before = xml.lengthAttr(node, "before");
		var after = xml.lengthAttr(node, "after");
		var line = xml.intAttr(node, "line", null);
		var lineRule = xml.attr(node, "lineRule");

		if (before) style["margin-top"] = before;
		if (after) style["margin-bottom"] = after;

		if (line !== null) {
			switch (lineRule) {
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
		xmlUtil.foreach(node, c => {
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
		switch (xml.attr(node, "hRule")) {
			case "exact":
				output["height"] = xml.lengthAttr(node, "val");
				break;

			case "atLeast":
			default:
				output["height"] = xml.lengthAttr(node, "val");
				// min-height doesn't work for tr
				//output["min-height"] = xml.sizeAttr(node, "val");  
				break;
		}
	}

	parseBorderProperties(node: Element, output: Record<string, string>) {
		xmlUtil.foreach(node, c => {
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

const knownColors = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'];

class xmlUtil {
	static foreach(node: Element, cb: (n: Element) => void) {
		for (var i = 0; i < node.childNodes.length; i++) {
			let n = node.childNodes[i];

			if (n.nodeType == Node.ELEMENT_NODE)
				cb(<Element>n);
		}
	}

	static colorAttr(node: Element, attrName: string, defValue: string = null, autoColor: string = 'black') {
		var v = xml.attr(node, attrName);

		if (v) {
			if (v == "auto") {
				return autoColor;
			} else if (knownColors.includes(v)) {
				return v;
			}

			return `#${v}`;
		}

		var themeColor = xml.attr(node, "themeColor");

		return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
	}

	static sizeValue(node: Element, type: LengthUsageType = LengthUsage.Dxa) {
		return convertLength(node.textContent, type);
	}
}

class values {
	static themeValue(c: Element, attr: string) {
		var val = xml.attr(c, attr);
		return val ? `var(--docx-${val}-font)` : null;
	}

	static valueOfSize(c: Element, attr: string) {
		var type = LengthUsage.Dxa;

		switch (xml.attr(c, "type")) {
			case "dxa": break;
			case "pct": type = LengthUsage.Percent; break;
			case "auto": return "auto";
		}

		return xml.lengthAttr(c, attr, type);
	}

	static valueOfMargin(c: Element) {
		return xml.lengthAttr(c, "w");
	}

	static valueOfBorder(c: Element) {
		var type = values.parseBorderType(xml.attr(c, "val"));

		if (type == "none")
			return "none";

		var color = xmlUtil.colorAttr(c, "color");
		var size = xml.lengthAttr(c, "sz", LengthUsage.Border);

		return `${size} ${type} ${color == "auto" ? autos.borderColor : color}`;
	}

	static parseBorderType(type: string) {
		switch (type) {
			case "single": return "solid";
			case "dashDotStroked": return "solid";
			case "dashed": return "dashed";
			case "dashSmallGap": return "dashed";
			case "dotDash": return "dotted";
			case "dotDotDash": return "dotted";
			case "dotted": return "dotted";
			case "double": return "double";
			case "doubleWave": return "double";
			case "inset": return "inset";
			case "nil": return "none";
			case "none": return "none";
			case "outset": return "outset";
			case "thick": return "solid";
			case "thickThinLargeGap": return "solid";
			case "thickThinMediumGap": return "solid";
			case "thickThinSmallGap": return "solid";
			case "thinThickLargeGap": return "solid";
			case "thinThickMediumGap": return "solid";
			case "thinThickSmallGap": return "solid";
			case "thinThickThinLargeGap": return "solid";
			case "thinThickThinMediumGap": return "solid";
			case "thinThickThinSmallGap": return "solid";
			case "threeDEmboss": return "solid";
			case "threeDEngrave": return "solid";
			case "triple": return "double";
			case "wave": return "solid";
		}

		return 'solid';
	}

	static valueOfTblLayout(c: Element) {
		var type = xml.attr(c, "val");
		return type == "fixed" ? "fixed" : "auto";
	}

	static classNameOfCnfStyle(c: Element) {
		const val = xml.attr(c, "val");
		const classes = [
			'first-row', 'last-row', 'first-col', 'last-col',
			'odd-col', 'even-col', 'odd-row', 'even-row',
			'ne-cell', 'nw-cell', 'se-cell', 'sw-cell'
		];

		return classes.filter((_, i) => val[i] == '1').join(' ');
	}

	static valueOfJc(c: Element) {
		var type = xml.attr(c, "val");

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

	static valueOfVertAlign(c: Element, asTagName: boolean = false) {
		var type = xml.attr(c, "val");

		switch (type) {
			case "subscript": return "sub";
			case "superscript": return asTagName ? "sup" : "super";
		}

		return asTagName ? null : type;
	}

	static valueOfTextAlignment(c: Element) {
		var type = xml.attr(c, "val");

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
		const val = xml.hexAttr(c, "val", 0);
		let className = "";

		if (xml.boolAttr(c, "firstRow") || (val & 0x0020)) className += " first-row";
		if (xml.boolAttr(c, "lastRow") || (val & 0x0040)) className += " last-row";
		if (xml.boolAttr(c, "firstColumn") || (val & 0x0080)) className += " first-col";
		if (xml.boolAttr(c, "lastColumn") || (val & 0x0100)) className += " last-col";
		if (xml.boolAttr(c, "noHBand") || (val & 0x0200)) className += " no-hband";
		if (xml.boolAttr(c, "noVBand") || (val & 0x0400)) className += " no-vband";

		return className.trim();
	}
}