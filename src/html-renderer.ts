import { WordDocument } from './word-document';
import { IDomNumbering, DocxContainer, DocxElement } from './document/dom';
import { Length, Underline } from './document/common';
import { Options } from './docx-preview';
import { WmlParagraph, ParagraphProperties } from './document/paragraph';
import { appendClass, keyBy, mergeDeep } from './utils';
import { updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { SectionProperties } from './document/section';
import { WmlRun, RunFonts, RunProperties, Shading } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { NumberingPartProperties } from './numbering/numbering';
import { Border } from './document/border';
import { TableColumn, WmlTable } from './document/table';
import { WmlTableRow } from './document/table-row';
import { WmlTableCell } from './document/table-cell';
import { WmlHyperlink } from './document/hyperlink';
import { WmlDrawing, DmlPicture } from './document/drawing';
import { WmlBreak, WmlLastRenderedPageBreak } from './document/breaks';
import { WmlSymbol, WmlTab, WmlText } from './document/text';
import { LineSpacing } from './document/line-spacing';
import { WmlStyle } from './styles/style';
import { WmlHeader } from './header/header';
import { WmlFooter } from './footer/footer';
import { WmlBody } from './document/document';

const knownColors = ['black','blue','cyan','darkBlue','darkCyan','darkGray','darkGreen','darkMagenta','darkRed','darkYellow','green','lightGray','magenta','none','red','white','yellow'];

export var autos = {
    shd: "white",
    color: "black",
    highlight: "transparent"
};

export class HtmlRenderer {

    inWrapper: boolean = true;
    className: string = "docx";
    document: WordDocument;
    options: Options;
    domStyleMap: Record<string, IDomStyle>;
    styleMap: Record<string, WmlStyle>;
    keepOrigin: boolean = false;
    renderHeaders: boolean = true;
    renderFooters: boolean = true;

    constructor(private htmlDocument: HTMLDocument) {
    }

    render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
        this.document = document;
        this.options = options;
        this.domStyleMap = null;

        styleContainer = styleContainer || bodyContainer;

        removeAllElements(styleContainer);
        removeAllElements(bodyContainer);

        appendComment(styleContainer, "docxjs library predefined styles");
        styleContainer.appendChild(this.renderDefaultStyle());
        
        if (document.stylesPart != null) {
            this.domStyleMap = this.processDomStyles(document.stylesPart.domStyles);
            this.styleMap = this.processStyles(document.stylesPart.styles);

            appendComment(styleContainer, "docx document styles");
            styleContainer.appendChild(this.renderStyles(document.stylesPart.domStyles));
        }

        if (document.numberingPart) {
            appendComment(styleContainer, "docx document numbering styles");
            styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
            //styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
        }

        if(!options.ignoreFonts && document.fontTablePart)
            this.renderFontTable(document.fontTablePart, styleContainer);

        var sectionElements = this.renderSections(document.documentPart.documentElement.body);

        if (this.inWrapper) {
            var wrapper = this.renderWrapper();
            appentElements(wrapper, sectionElements);
            bodyContainer.appendChild(wrapper);
        }
        else {
            appentElements(bodyContainer, sectionElements);
        }
    }

    renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
        for(let f of fontsPart.fonts.filter(x => x.refId)) {
            this.document.loadFont(f.refId, f.fontKey).then(fontData => {
                var cssTest = `@font-face {
                    font-family: "${f.name}";
                    src: url(${fontData});
                }`;

                appendComment(styleContainer, `Font ${f.name}`);
                styleContainer.appendChild(createStyleElement(cssTest));
            });
        }
    }

    processClassName(className: string) {
        if (!className)
            return this.className;

        return `${this.className}_${className}`;
    }

    processStyles(styles: WmlStyle[]) {
        const styleMap = keyBy(styles, s => s.id);

        for(let style of styles.filter(s => s.basedOn)) {
            const baseStyle = styleMap[style.basedOn];
        
            if(baseStyle) {
                style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
                style.runProps = mergeDeep(style.runProps, baseStyle.runProps);
            } else if (this.options.debug) {
                console.warn(`Can't find base style ${style.basedOn}`);
            }
        }

        return styleMap;
    }

    processDomStyles(styles: IDomStyle[]) {
        var domStylesMap: Record<string, IDomStyle> = {};

        for (let style of styles.filter(x => x.id != null)) {
            domStylesMap[style.id] = style;
        }

        for (let style of styles.filter(x => x.basedOn)) {
            var baseStyle = domStylesMap[style.basedOn];

            if (baseStyle) {
                for (let styleValues of style.styles) {
                    var baseValues = baseStyle.styles.filter(x => x.target == styleValues.target);

                    if (baseValues && baseValues.length > 0)
                        this.copyStyleProperties(baseValues[0].values, styleValues.values);
                }
            }
            else if (this.options.debug)
                console.warn(`Can't find base style ${style.basedOn}`);
        }

        for (let style of styles) {
            style.cssName = this.processClassName(this.escapeClassName(style.id));
        }

        return domStylesMap;
    }

    processElement(element: DocxElement) {
        if ("children" in element) {
            for (var e of (element as DocxContainer).children) {
                e.className = this.processClassName(e.className);
                e.parent = element;

                if (e instanceof WmlTable) {
                    this.processTable(e);
                }
                else {
                    this.processElement(e);
                }
            }
        }
    }

    processTable(table: WmlTable) {
        for (var r of table.children) {
            for (var c of (r as DocxContainer).children) {
                c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
                    "border-left", "border-right", "border-top", "border-bottom",
                    "padding-left", "padding-right", "padding-top", "padding-bottom"
                ]);

                this.processElement(c);
            }
        }
    }

    copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
        if (!input)
            return output;

        if (output == null) output = {};
        if (attrs == null) attrs = Object.getOwnPropertyNames(input);

        for (var key of attrs) {
            if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                output[key] = input[key];
        }

        return output;
    }

    private createElement(tagName, props = undefined) {
        return Object.assign(this.htmlDocument.createElement(tagName), props);
    }

    private renderContainer(elem: DocxContainer, tagName: string): HTMLElement {
        const result = this.createElement(tagName);
        this.renderElements(elem.children, elem, result);
        return result;
    }

    createSection(className: string, props: SectionProperties) {
        var elem = this.createElement("section", { className });
        
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
                elem.style.columnCount = `${props.columns.numberOfColumns}`;
                elem.style.columnGap = this.renderLength(props.columns.space);

                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }

        return elem;
    }

    renderSections(document: WmlBody): HTMLElement[] {
        const result = [];

        this.processElement(document);

        for(let section of this.splitBySection(document.children)) {
            const sectionProps = section.sectProps || document.sectionProps;
            var sectionElement = this.createSection(this.className, sectionProps);

            // if(this.renderHeaders && sectionProps.headerRefs) {
            //     for(const headerRef of sectionProps.headerRefs) {
            //     }
            // }

            this.renderElements(section.elements, document, sectionElement);

            // if(this.renderFooters && sectionProps.footerRefs) {
            //     for(const headerRef of sectionProps.footerRefs) {
            //         const partPath = this.document.getPathById(this.document.documentPart, headerRef.id);
            //         const part = this.document.partsMap[partPath] as FooterPart;

            //         this.processElement(part.footerElement);
            //         sectionElement.appendChild(this.renderElement(part.footerElement, null))
            //     }
            // }

            result.push(sectionElement);
        }

        return result;
    }

    isPageBreakElement(elem: DocxElement): boolean {
        if (elem instanceof WmlLastRenderedPageBreak)
            return !this.options.ignoreLastRenderedPageBreak;

        return elem instanceof WmlBreak && elem.type === "page";  
    }

    splitBySection(elements: DocxElement[]): { sectProps: SectionProperties, elements: DocxElement[] }[] {
        var current = { sectProps: null, elements: [] };
        var result = [current];

        for(let elem of elements) {
            if (elem instanceof WmlParagraph) {
                const styleName = elem.props.styleId;
                const s = this.styleMap && styleName ? this.styleMap[styleName] : null;
            
                if(s?.paragraphProps?.pageBreakBefore) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }
            }

            current.elements.push(elem);

            if(elem instanceof WmlParagraph)
            {
                const p = elem as WmlParagraph;

                var sectProps = p.props.sectionProps;
                var pBreakIndex = -1;
                var rBreakIndex = -1;
                
                if(this.options.breakPages && p.children.length > 0) {
                    pBreakIndex = p.children.findIndex((r: DocxContainer) => {
                        rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
                        return rBreakIndex != -1;
                    });
                }
    
                if(sectProps || pBreakIndex != -1) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }

                if(pBreakIndex != -1) {
                    let breakRun = p.children[pBreakIndex] as WmlRun;
                    let splitRun = rBreakIndex < breakRun.children.length - 1;

                    if(pBreakIndex < p.children.length - 1 || splitRun) {
                        var children = elem.children;
                        var newParagraph = Object.assign(new WmlParagraph(), elem, { children: children.slice(pBreakIndex) });
                        elem.children = children.slice(0, pBreakIndex);
                        current.elements.push(newParagraph);

                        if(splitRun) {
                            let runChildren = breakRun.children;
                            let newRun =  Object.assign(new WmlRun(), breakRun, { children: runChildren.slice(0, rBreakIndex) });
                            elem.children.push(newRun);
                            breakRun.children = runChildren.slice(rBreakIndex);
                        }
                    }
                }
            }
        }

        let currentSectProps = null;

        for (let i = result.length - 1; i >= 0; i--) {
            if (result[i].sectProps == null) {
                result[i].sectProps = currentSectProps;
            } else {
                currentSectProps = result[i].sectProps
            }
        }

        return result;
    }

    renderLength(l: Length): string {
        return l ? `${l.value}${l.type}` : null;
    }

    renderColor(c: string, autoColor: string = 'black'): string {
        if (/[a-f0-9]{6}/i.test(c))
            return `#${c}`;

        return c === 'auto' ? autoColor : c;
    }

    renderWrapper() {
        var wrapper = document.createElement("div");

        wrapper.className = `${this.className}-wrapper`

        return wrapper;
    }

    renderDefaultStyle() {
        var styleText = `.${this.className}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
                .${this.className}-wrapper section.${this.className} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
                .${this.className} { color: black; }
                section.${this.className} { box-sizing: border-box; }
                .${this.className} table { border-collapse: collapse; }
                .${this.className} table td, .${this.className} table th { vertical-align: top; }
                .${this.className} p { margin: 0pt; }`;

        return createStyleElement(styleText);
    }

    // renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
    //     let css = "";
    //     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
    //     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
    //     const topCounters = [];

    //     for(let num of numberingPart.numberings) {
    //         const absNum = numberingMap[num.abstractId];

    //         for(let lvl of absNum.levels) {
    //             const className = this.numberingClass(num.id, lvl.level);
    //             let listStyleType = "none";

    //             if(lvl.text && lvl.format == 'decimal') {
    //                 const counter = this.numberingCounter(num.id, lvl.level);

    //                 if (lvl.level > 0) {
    //                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
    //                         "counter-reset": counter
    //                     });
    //                 } else {
    //                     topCounters.push(counter);
    //                 }
                    
    //                 css += this.styleToString(`p.${className}:before`, {
    //                     "content": this.levelTextToContent(lvl.text, num.id),
    //                     "counter-increment": counter
    //                 });
    //             } else if(lvl.bulletPictureId) {
    //                 let pict = bulletMap[lvl.bulletPictureId];
    //                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();

    //                 css += this.styleToString(`p.${className}:before`, {
    //                     "content": "' '",
    //                     "display": "inline-block",
    //                     "background": `var(${variable})`
    //                 }, pict.style);
    
    //                 this.document.loadNumberingImage(pict.referenceId).then(data => {
    //                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
    //                     container.appendChild(createStyleElement(text));
    //                 });
    //             } else {
    //                 listStyleType = this.numFormatToCssValue(lvl.format);
    //             }

    //             css += this.styleToString(`p.${className}`, {
    //                 "display": "list-item",
    //                 "list-style-position": "inside",
    //                 "list-style-type": listStyleType,
    //                 //TODO
    //                 //...num.style
    //             });
    //         }
    //     }

    //     if (topCounters.length > 0) {
    //         css += this.styleToString(`.${this.className}-wrapper`, {
    //             "counter-reset": topCounters.join(" ")
    //         });
    //     }

    //     return createStyleElement(css);
    // }

    renderNumbering(styles: IDomNumbering[], styleContainer: HTMLElement) {
        var styleText = "";
        var rootCounters = [];

        for (var num of styles) {
            var selector = `p.${this.numberingClass(num.id, num.level)}`;
            var listStyleType = "none";

            if (num.levelText && (num.format == "decimal" || num.format == "lowerLetter" || num.format == "lowerRoman")) {
                let counter = this.numberingCounter(num.id, num.level);

                if (num.level > 0) {
                    styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
                        "counter-reset": counter
                    });
                }
                else {
                    rootCounters.push(counter);
                }

                styleText += this.styleToString(`${selector}:before`, {
                    "content": this.levelTextToContent(num.levelText, num.id, this.numFormatToCssValue(num.format)),
                    "counter-increment": counter
                });
            }
            else if (num.bullet) {
                let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

                styleText += this.styleToString(`${selector}:before`, {
                    "content": "' '",
                    "display": "inline-block",
                    "background": `var(${valiable})`
                }, num.bullet.style);

                this.document.loadNumberingImage(num.bullet.src).then(data => {
                    var text = `.${this.className}-wrapper { ${valiable}: url(${data}) }`;
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else {
                listStyleType = this.numFormatToCssValue(num.format);
            }

            styleText += this.styleToString(selector, {
                "display": "list-item",
                "list-style-position": "inside",
                "list-style-type": listStyleType,
                ...num.style
            });
        }

        if (rootCounters.length > 0) {
            styleText += this.styleToString(`.${this.className}-wrapper`, {
                "counter-reset": rootCounters.join(" ")
            });
        }

        return createStyleElement(styleText);
    }

    renderStyles(styles: IDomStyle[]): HTMLElement {
        var styleText = "";
        var stylesMap = this.domStyleMap;

        for (let style of styles) {
            var subStyles =  style.styles;

            if(style.linked) {
                var linkedStyle = style.linked && stylesMap[style.linked];

                if (linkedStyle)
                    subStyles = subStyles.concat(linkedStyle.styles);
                else if(this.options.debug)
                    console.warn(`Can't find linked style ${style.linked}`);
            }

            for (var subStyle of subStyles) {
                var selector = "";

                if (style.target == subStyle.target)
                    selector += `${style.target}.${style.cssName}`;
                else if (style.target)
                    selector += `${style.target}.${style.cssName} ${subStyle.target}`;
                else
                    selector += `.${style.cssName} ${subStyle.target}`;

                if (style.isDefault && style.target)
                    selector = `.${this.className} ${style.target}, ` + selector;

                if (style.paragraphProps && subStyle.target == "p") {
                    this.renderParagraphProperties(subStyle.values, style.paragraphProps);
                }

                styleText += this.styleToString(selector, subStyle.values);
            }
        }

        return createStyleElement(styleText);
    }

    renderElement(elem: DocxElement, parent: DocxElement): Node {
        if (elem instanceof WmlParagraph) {
            return this.renderParagraph(elem);
        } else if (elem instanceof WmlBookmarkStart) {
            return this.renderBookmarkStart(elem);
        } else if (elem instanceof WmlRun) {
            return this.renderRun(elem);
        } else if (elem instanceof WmlText) {
            return this.renderText(elem);
        } else if (elem instanceof WmlSymbol) {
            return this.renderSymbol(elem);
        } else if (elem instanceof WmlTab) {
            return this.renderTab(elem);
        } else if (elem instanceof WmlTable) {
            return this.renderTable(elem);
        } else if (elem instanceof WmlTableRow) {
            return this.renderTableRow(elem);
        } else if (elem instanceof WmlTableCell) {
            return this.renderTableCell(elem);
        } else if (elem instanceof WmlHyperlink) {
            return this.renderHyperlink(elem);
        } else if (elem instanceof WmlDrawing) {
            return this.renderDrawing(elem);
        } else if (elem instanceof DmlPicture) {
            return this.renderImage(elem);
        } else if (elem instanceof WmlHeader) {
            return this.renderHeader(elem);
        } else if (elem instanceof WmlFooter) {
            return this.renderFooter(elem);
        }

        return null;
    }

    renderChildren(elem: DocxContainer, into?: HTMLElement): Node[] {
        return this.renderElements(elem.children, elem, into);
    }

    renderElements(elems: DocxElement[], parent: DocxElement, into?: HTMLElement): Node[] {
        if(elems == null)
            return null;

        var result = elems.map(e => {
            let n = this.renderElement(e, parent);

            if(n && this.keepOrigin)
                (n as any).$$docxElement = e;

            return n;
        }).filter(e => e != null);

        if(into)
            for(let c of result)
                into.appendChild(c);

        return result;
    }

    renderParagraph(elem: WmlParagraph) {
        var result = this.renderContainer(elem, "p");

        this.renderClass(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        const style = elem.props.styleId && this.styleMap?.[elem.props.styleId]; 
        const numbering = elem.props.numbering ?? style?.paragraphProps?.numbering;

        if (numbering) {
            var numberingClass = this.numberingClass(numbering.id, numbering.level ?? 0);
            result.className = appendClass(result.className, numberingClass);
        }

        if (elem.props.styleId) {
            var styleClassName = this.processClassName(this.escapeClassName(elem.props.styleId));
            result.className = appendClass(result.className, styleClassName);
        }

        return result;
    }

    renderParagraphProperties(style: any, props: ParagraphProperties) {
        for (const p in props) {
            const v = props[p];

            switch(p as keyof(ParagraphProperties)) {
                case "lineSpacing":
                    this.renderLineSpacing(style, v);
                    break;
            }
        }
    }

    renderLineSpacing(style: any, spacing: LineSpacing) {   
        //TODO
    }

    renderRunProperties(style: any, props: RunProperties) {
        for (const p in props) {
            const v = props[p];

            switch (p as keyof(RunProperties)) {
                case 'highlight':
                    style['background'] = this.renderColor(v);
                    break;

                case 'shading':
                    style['background'] = this.renderShading(v);
                    break;

                case 'border':
                    style['border'] = this.renderBorder(v);
                    break;

                case 'color':
                    style["color"] = this.renderColor(v);
                    break;

                case 'fontSize':
                    style["font-size"] = this.renderLength(v);
                    break;

                case 'bold':
                    style["font-weight"] = v ? 'bold' : 'normal';
                    break;

                case 'italics':
                    style["font-style"] = v ? 'italic' : 'normal';
                    break;

                case 'smallCaps':
                    style["font-size"] = v ? 'smaller' : 'none';
                case 'caps':
                    style["text-transform"] = v ? 'uppercase' : 'none';
                    break;
                    
                case 'strike':
                    style["text-decoration"] = v ? 'line-through' : 'none';
                    break;

                case 'fonts':
                    style["font-family"] = this.renderRunFonts(v);
                    break;
    
                case 'underline':
                    this.renderUnderline(style, v);
                    break;
                
                case 'verticalAlignment':
                    this.renderRunVerticalAlignment(style, v);
                    break;
            }
        }
    }

    renderRunVerticalAlignment(style: any, align: string) {
        switch(align) {
            case 'subscript': 
                style['vertical-align'] = 'sub';
                style['font-size'] = 'small';
                break;

            case 'superscript': 
                style['vertical-align'] = 'super';
                style['font-size'] = 'small';
                break;
        }
    }

    renderRunFonts(fonts: RunFonts) {
        return [fonts.ascii, fonts.hAscii, fonts.cs, fonts.eastAsia].filter(x => x).map(x => `'${x}'`).join(',');
    }

    renderBorder(border: Border) {
        if (border.type == 'nil')
            return 'none';

        return `${this.renderLength(border.size)} solid ${this.renderColor(border.color)}`;
    }
    
    renderShading(shading: Shading) {
        if (shading.type == 'clear')
            return this.renderColor(shading.background, autos.shd);
        
        return this.renderColor(shading.background, autos.shd);
    }
    
    renderUnderline(style: any, underline: Underline) {
        if (underline.type == null || underline.type == "none")
            return;

        switch (underline.type) {
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

        if (underline.color)
            style["text-decoration-color"] = this.renderColor(underline.color);
    }

    renderHyperlink(elem: WmlHyperlink) {
        var result = this.createElement("a");

        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        if (elem.anchor)
            result.href = elem.anchor;

        return result;
    }

    renderDrawing(elem: WmlDrawing) {
        var result = this.createElement("div");

        result.style.display = "inline-block";
        result.style.position = "relative";
        result.style.textIndent = "0px";

        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        return result;
    }

    renderImage(elem: DmlPicture) {
        let result = this.createElement("img");

        this.renderStyleValues(elem.cssStyle, result);

        if (this.document) {
            this.document.loadDocumentImage(elem.resourceId).then(x => {
                result.src = x;
            });
        }

        return result;
    }

    renderHeader(elem: WmlHeader) {
        return this.renderContainer(elem, "header");
    }

    renderFooter(elem: WmlHeader) {
        return this.renderContainer(elem, "footer");
    }

    renderText(elem: WmlText) {
        return this.htmlDocument.createTextNode(elem.text);
    }

    renderSymbol(elem: WmlSymbol) {
        var span = this.createElement("span");
        span.style.fontFamily = elem.font;
        span.innerHTML = `&#x${elem.char};`
        return span;
    }

    renderTab(elem: WmlTab) {
        var tabSpan = this.createElement("span");
     
        tabSpan.innerHTML = "&emsp;";//"&nbsp;";

        if(this.options.experimental) {
            setTimeout(() => {
                var paragraph = findParent<WmlParagraph>(elem, WmlParagraph);
                
                if(paragraph.props.tabs == null)
                    return;

                paragraph.props.tabs.sort((a, b) => a.position.value - b.position.value);
                tabSpan.style.display = "inline-block";
                updateTabStop(tabSpan, paragraph.props.tabs);
            }, 0);
        }

        return tabSpan;
    }

    renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
        var result = this.createElement("span");
        result.id = elem.name;
        return result;
    }

    renderRun(elem: WmlRun) {
        var result = this.createElement("span");

        if(elem.id)
            result.id = elem.id;

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        //this.renderStyleValues(elem.cssStyle, result);
        this.renderRunProperties(result.style, elem.props);

        return result;
    }

    renderTable(elem: WmlTable) {
        let result = this.createElement("table");

        if (elem.columns)
            result.appendChild(this.renderTableColumns(elem.columns));

        if (elem.props?.caption) {
            result.appendChild(this.createElement("caption", { textContent: elem.props?.caption }));
        }

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        return result;
    }

    renderTableColumns(columns: TableColumn[]) {
        const result = this.createElement("colGroup");

        for (let col of columns) {
            let colElem = this.createElement("col");
            colElem.style.width = this.renderLength(col.width);
            result.appendChild(colElem);
        }

        return result;
    }

    renderTableRow(elem: WmlTableRow) {
        let result = this.createElement("tr");

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        return result;
    }

    renderTableCell(elem: WmlTableCell) {
        let result = this.createElement("td");

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        if (elem.span) result.colSpan = elem.span;

        return result;
    }

    renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
        if (style == null)
            return;

        for (let key of Object.getOwnPropertyNames(style)) {
            ouput.style[key] = style[key];
        }
    }

    renderClass(input: DocxElement, ouput: HTMLElement) {
        if (input.className)
            ouput.className = input.className;
    }

    numberingClass(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
        let result = `${selectors} {\r\n`;

        for (const key in values) {
            result += `  ${key}: ${values[key]};\r\n`;
        }

        if (cssText) 
            result += cssText;

        return result + "}\r\n";
    }

    numberingCounter(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    levelTextToContent(text: string, id: string, numformat: string) {
        var result = text.replace(/%\d*/g, s => {
            let lvl = parseInt(s.substring(1), 10) - 1;
            return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
        });

        return '"' + result + '"';
    }

    numFormatToCssValue(format: string) {
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
    }

    escapeClassName(className: string) {
        return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and');
    }
}

function appentElements(container: HTMLElement, children: HTMLElement[]) {
    for (let c of children)
        container.appendChild(c);
}

function removeAllElements(elem: HTMLElement) {
    elem.innerHTML = '';
}

function createStyleElement(cssText: string) {
    var styleElement = document.createElement("style");
    styleElement.innerHTML = cssText;
    return styleElement;
}

function appendComment(elem: HTMLElement, comment: string) {
    elem.appendChild(document.createComment(comment));
}

function findParent<T extends DocxElement>(elem: DocxElement, type: any): T {
    var parent = elem.parent;

    while (parent != null && !(parent instanceof type))
        parent = parent.parent;
    
    return <T>parent;
}