import { WordDocument } from './word-document';
import { DomType, IDomTable, IDomNumbering, 
    IDomHyperlink, IDomImage, OpenXmlElement, IDomTableColumn, IDomTableCell, TextElement, SymbolElement, BreakElement } from './dom/dom';
import { Length, CommonProperties } from './dom/common';
import { Options } from './docx-preview';
import { DocumentElement } from './dom/document';
import { ParagraphElement } from './dom/paragraph';
import { appendClass, keyBy } from './utils';
import { updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { SectionProperties } from './dom/section';
import { RunElement, RunProperties } from './dom/run';
import { BookmarkStartElement } from './dom/bookmark';
import { IDomStyle } from './dom/style';
import { NumberingPartProperties } from './numbering/numbering';

export class HtmlRenderer {

    inWrapper: boolean = true;
    className: string = "docx";
    document: WordDocument;
    options: Options;

    constructor(public htmlDocument: HTMLDocument) {
    }

    render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
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
            styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
            //styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
        }

        if(!options.ignoreFonts && document.fontTablePart)
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

    processStyles(styles: IDomStyle[]) {
        var stylesMap: Record<string, IDomStyle> = {};

        for (let style of styles.filter(x => x.id != null)) {
            stylesMap[style.id] = style;
        }

        for (let style of styles.filter(x => x.basedOn)) {
            var baseStyle = stylesMap[style.basedOn];

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

        return stylesMap;
    }

    processElement(element: OpenXmlElement) {
        if (element.children) {
            for (var e of element.children) {
                e.className = this.processClassName(e.className);
                e.parent = element;

                if (e.type == DomType.Table) {
                    this.processTable(e);
                }
                else {
                    this.processElement(e);
                }
            }
        }
    }

    processTable(table: IDomTable) {
        for (var r of table.children) {
            for (var c of r.children) {
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

    createSection(className: string, props: SectionProperties) {
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
                elem.style.columnCount = `${props.columns.numberOfColumns}`;
                elem.style.columnGap = this.renderLength(props.columns.space);

                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }

        return elem;
    }

    renderSections(document: DocumentElement): HTMLElement[] {
        var result = [];

        this.processElement(document);

        for(let section of this.splitBySection(document.children)) {
            var sectionElement = this.createSection(this.className, section.sectProps || document.props);
            this.renderElements(section.elements, document, sectionElement);
            result.push(sectionElement);
        }

        return result;
    }

    splitBySection(elements: OpenXmlElement[]): { sectProps: SectionProperties, elements: OpenXmlElement[] }[] {
        var current = { sectProps: null, elements: [] };
        var result = [current];
        var styles = this.document.stylesPart?.styles;
        var styleMap = styles ? keyBy(styles, x => x.id) : null;

        for(let elem of elements) {
            if(elem.type == DomType.Paragraph) {
                const styleName = (elem as ParagraphElement).styleName;
                const s = styleMap && styleName ? styleMap[styleName] : null;
            
                if(s?.paragraphProps?.pageBreakBefore) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }
            }

            current.elements.push(elem);

            if(elem.type == DomType.Paragraph)
            {
                const p = elem as ParagraphElement;

                var sectProps = p.sectionProps;
                var pBreakIndex = -1;
                var rBreakIndex = -1;
                
                if(this.options.breakPages && p.children) {
                    pBreakIndex = p.children.findIndex(r => {
                        rBreakIndex = r.children?.findIndex(t => (t as BreakElement).break == "page") ?? -1;
                        return rBreakIndex != -1;
                    });
                }
    
                if(sectProps || pBreakIndex != -1) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }

                if(pBreakIndex != -1) {
                    let breakRun = p.children[pBreakIndex];
                    let splitRun = rBreakIndex < breakRun.children.length - 1;

                    if(pBreakIndex < p.children.length - 1 || splitRun) {
                        var children = elem.children;
                        var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
                        elem.children = children.slice(0, pBreakIndex);
                        current.elements.push(newParagraph);

                        if(splitRun) {
                            let runChildren = breakRun.children;
                            let newRun =  { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
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

            if (num.levelText && num.format == "decimal") {
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
                    "content": this.levelTextToContent(num.levelText, num.id),
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
        var stylesMap = this.processStyles(styles);

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

                styleText += this.styleToString(selector, subStyle.values);
            }
        }

        return createStyleElement(styleText);
    }

    renderElement(elem: OpenXmlElement, parent: OpenXmlElement): Node {
        switch (elem.type) {
            case DomType.Paragraph:
                return this.renderParagraph(<ParagraphElement>elem);

            case DomType.BookmarkStart:
                return this.renderBookmarkStart(<BookmarkStartElement>elem);

            case DomType.BookmarkEnd:
                return null;
    
            case DomType.Run:
                return this.renderRun(<RunElement>elem);

            case DomType.Table:
                return this.renderTable(elem);

            case DomType.Row:
                return this.renderTableRow(elem);

            case DomType.Cell:
                return this.renderTableCell(elem);

            case DomType.Hyperlink:
                return this.renderHyperlink(elem);

            case DomType.Drawing:
                return this.renderDrawing(<IDomImage>elem);

            case DomType.Image:
                return this.renderImage(<IDomImage>elem);
            
            case DomType.Text:
                return this.renderText(<TextElement>elem);

            case DomType.Tab:
                return this.renderTab(elem);
            
            case DomType.Symbol:
                return this.renderSymbol(<SymbolElement>elem);
        }

        return null;
    }

    renderChildren(elem: OpenXmlElement, into?: HTMLElement): Node[] {
        return this.renderElements(elem.children, elem, into);
    }

    renderElements(elems: OpenXmlElement[], parent: OpenXmlElement, into?: HTMLElement): Node[] {
        if(elems == null)
            return null;

        var result = elems.map(e => this.renderElement(e, parent)).filter(e => e != null);

        if(into)
            for(let c of result)
                into.appendChild(c);

        return result;
    }

    renderParagraph(elem: ParagraphElement) {
        var result = this.htmlDocument.createElement("p");

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        this.renderCommonProeprties(result.style, elem);

        if (elem.numbering) {
            var numberingClass = this.numberingClass(elem.numbering.id, elem.numbering.level);
            result.className = appendClass(result.className, numberingClass);
        }

        if (elem.styleName) {
            var styleClassName = this.processClassName(this.escapeClassName(elem.styleName));
            result.className = appendClass(result.className, styleClassName);
        }

        return result;
    }

    renderRunProperties(style: any, props: RunProperties) {
        this.renderCommonProeprties(style, props);
    }

    renderCommonProeprties(style: any, props: CommonProperties){
        if(props == null)
            return;

        if(props.color) {
            style["color"] = props.color;
        }

        if (props.fontSize) {
            style["font-size"] = this.renderLength(props.fontSize);
        }
    }

    renderHyperlink(elem: IDomHyperlink) {
        var result = this.htmlDocument.createElement("a");

        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        if (elem.href)
            result.href = elem.href

        return result;
    }

    renderDrawing(elem: IDomImage) {
        var result = this.htmlDocument.createElement("div");

        result.style.display = "inline-block";
        result.style.position = "relative";
        result.style.textIndent = "0px";

        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        return result;
    }

    renderImage(elem: IDomImage) {
        let result = this.htmlDocument.createElement("img");

        this.renderStyleValues(elem.cssStyle, result);

        if (this.document) {
            this.document.loadDocumentImage(elem.src).then(x => {
                result.src = x;
            });
        }

        return result;
    }

    renderText(elem: TextElement) {
        return this.htmlDocument.createTextNode(elem.text);
    }

    renderSymbol(elem: SymbolElement) {
        var span = this.htmlDocument.createElement("span");
        span.style.fontFamily = elem.font;
        span.innerHTML = `&#x${elem.char};`
        return span;
    }

    renderTab(elem: OpenXmlElement) {
        var tabSpan = this.htmlDocument.createElement("span");
     
        tabSpan.innerHTML = "&emsp;";//"&nbsp;";

        if(this.options.experimental) {
            setTimeout(() => {
                var paragraph = findParent<ParagraphElement>(elem, DomType.Paragraph);
                
                if(paragraph.tabs == null)
                    return;

                paragraph.tabs.sort((a, b) => a.position.value - b.position.value);
                tabSpan.style.display = "inline-block";
                updateTabStop(tabSpan, paragraph.tabs);
            }, 0);
        }

        return tabSpan;
    }

    renderBookmarkStart(elem: BookmarkStartElement): HTMLElement {
        var result = this.htmlDocument.createElement("span");
        result.id = elem.name;
        return result;
    }

    renderRun(elem: RunElement) {
        if (elem.break)
            return elem.break == "page" ? null : this.htmlDocument.createElement("br");
        
        if (elem.fldCharType || elem.instrText)
            return null;

        var result = this.htmlDocument.createElement("span");

        if(elem.id)
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
    }

    renderTable(elem: IDomTable) {
        let result = this.htmlDocument.createElement("table");

        if (elem.columns)
            result.appendChild(this.renderTableColumns(elem.columns));

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        return result;
    }

    renderTableColumns(columns: IDomTableColumn[]) {
        let result = this.htmlDocument.createElement("colGroup");

        for (let col of columns) {
            let colElem = this.htmlDocument.createElement("col");

            if (col.width)
                colElem.style.width = `${col.width}px`;

            result.appendChild(colElem);
        }

        return result;
    }

    renderTableRow(elem: OpenXmlElement) {
        let result = this.htmlDocument.createElement("tr");

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        return result;
    }

    renderTableCell(elem: IDomTableCell) {
        let result = this.htmlDocument.createElement("td");

        this.renderClass(elem, result);
        this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        if (elem.span) result.colSpan = elem.span;

        return result;
    }

    renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
        if (style == null)
            return;

        for (let key in style) {
            if (style.hasOwnProperty(key)) {
                ouput.style[key] = style[key];
            }
        }
    }

    renderClass(input: OpenXmlElement, ouput: HTMLElement) {
        if (input.className)
            ouput.className = input.className;
    }

    numberingClass(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
        let result = selectors + " {\r\n";

        for (const key in values) {
            result += `  ${key}: ${values[key]};\r\n`;
        }

        if (cssText)
            result += ";" + cssText;

        return result + "}\r\n";
    }

    numberingCounter(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    levelTextToContent(text: string, id: string) {
        var result = text.replace(/%\d*/g, s => {
            let lvl = parseInt(s.substring(1), 10) - 1;
            return `"counter(${this.numberingCounter(id, lvl)})"`;
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
    while (elem.firstChild) {
        elem.removeChild(elem.firstChild);
    }
}

function createStyleElement(cssText: string) {
    var styleElement = document.createElement("style");
    styleElement.innerHTML = cssText;
    return styleElement;
}

function appendComment(elem: HTMLElement, comment: string) {
    elem.appendChild(document.createComment(comment));
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
    var parent = elem.parent;

    while (parent != null && parent.type != type)
        parent = parent.parent;
    
    return <T>parent;
}