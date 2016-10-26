namespace docx {
    export class HtmlRenderer {

        private digitTest = /^[0-9]/.test;

        constructor(public htmlDocument: HTMLDocument) {
        }

        processClassName(className){
            if(!className)
                return null;
                
            return "docx_" + className;            
        }

        processStyles(styles: IDomStyle[]) {
            var stylesMap = {};

            for(let style of styles){
                style.id = this.processClassName(style.id);
                style.basedOn = this.processClassName(style.basedOn);

                stylesMap[style.id] = style;
            }

            for(let style of styles){
                if(style.basedOn){
                    var baseStyle = stylesMap[style.basedOn];

                    for(let styleValues of style.styles){
                        var baseValues = baseStyle.styles.filter(x => x.target == styleValues.target);

                        if(baseValues && baseValues.length > 0)
                             this.copyStyleProperties(baseValues[0].values, styleValues.values);
                    }
                }
            }
        }

        processElement(element: IDomDocument) {
            if (element.children) {
                for (var e of element.children) {
                    e.className = this.processClassName(e.className);

                    if (e.domType == DomType.Table) {
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
                    c.style = this.copyStyleProperties(table.cellStyle, c.style, [
                        "border-left", "border-right", "border-top", "border-bottom",
                        "padding-left", "padding-right", "padding-top", "padding-bottom"
                    ]);

                    this.processElement(c);
                }
            }
        }

	    copyStyleProperties(input: IDomStyleValues, output: IDomStyleValues, attrs: string[] = null): IDomStyleValues {
            if(!input)
                return output;

            if(output == null) output = {};
            if(attrs == null) attrs = Object.getOwnPropertyNames(input);

            for (var key of attrs) {
                if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                    output[key] = input[key];
            }

            return output;
        }

        renderDocument(document: IDomDocument): HTMLElement {
            var bodyElement = this.htmlDocument.createElement("section");

            bodyElement.className = "docx";

            this.processElement(document);
            this.renderChildren(document, bodyElement);

            this.renderStyleValues(document.style, bodyElement);

            return bodyElement;
        }

        renderWrapper(){
            var wrapper = document.createElement("div");

            wrapper.className = "docx-wrapper"

            return wrapper;
        }

        renderDefaultStyle(){
            var styleElement = document.createElement("style");

            styleElement.type = "text/css";
            styleElement.innerHTML = ".docx-wrapper { background: gray; padding: 30px; display: flex; justify-content: center; }\r\n" 
                + ".docx-wrapper section.docx {background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); }\r\n"
                + ".docx { color: black; }\r\n"
                + "section.docx { box-sizing: border-box; }\r\n"
                + ".docx table { border-collapse: collapse; }\r\n"
                + ".docx table td, .docx table th { vertical-align: top; }\r\n"
                + ".docx p { margin: 0pt; }";

            return styleElement;
        }

        renderStyles(styles: IDomStyle[]): HTMLElement {
            var styleElement = document.createElement("style");
            var styleText = "";

            styleElement.type = "text/css";

            this.processStyles(styles);

            for (let style of styles) {
                for (var subStyle of style.styles) {
                    if (style.isDefault)
                        styleText += style.target + ", ";

                    if (style.target == subStyle.target)
                        styleText += style.target + "." + style.id + "{\r\n";
                    else 
                        styleText += style.target + "." + style.id + " " + subStyle.target + "{\r\n";

                    for (var key in subStyle.values) {
                        styleText += key + ": " + subStyle.values[key] + ";\r\n";
                    }

                    styleText += "}\r\n";
                }
            }

            styleElement.innerHTML = styleText;

            return styleElement;
        }

        renderElement(elem): HTMLElement {
            switch (elem.domType) {
                case DomType.Paragraph:
                    return this.renderParagraph(elem);

                case DomType.Run:
                    return this.renderRun(<IDomRun>elem);

                case DomType.Table:
                    return this.renderTable(elem);

                case DomType.Row:
                    return this.renderTableRow(elem);

                case DomType.Cell:
                    return this.renderTableCell(elem);
            }

            return null;
        }

        renderChildren(elem: IDomElement, into?: HTMLElement): HTMLElement[] {
            var result: HTMLElement[] = null;

            if (elem.children != null)
                result = elem.children.map(x => this.renderElement(x)).filter(x => x != null);

            if (into && result)
                result.forEach(x => into.appendChild(x));

            return result;
        }

        renderParagraph(elem: IDomParagraph) {
            var result = this.htmlDocument.createElement("p");

            this.renderClass(elem, result);
            this.renderChildren(elem, result);
            this.renderStyleValues(elem.style, result);

            return result;
        }

        renderRun(elem: IDomRun) {

            if (elem.break)
                return this.htmlDocument.createElement(elem.break == "page" ? "hr" : "br");

            var result = this.htmlDocument.createElement("span");

            this.renderStyleValues(elem.style, result);

            result.textContent = elem.text;

            return result;
        }

        renderTable(elem: IDomTable) {
            var result = this.htmlDocument.createElement("table");

            this.renderClass(elem, result);
            this.renderChildren(elem, result);
            this.renderStyleValues(elem.style, result);

            if(elem.columns)
                result.appendChild(this.renderTableColumns(elem.columns));

            return result;
        }

        renderTableColumns(columns: IDomTableColumn[]) {
            var result = this.htmlDocument.createElement("colGroup");

            for(let col of columns) {
                var colElem = this.htmlDocument.createElement("col");

                if(col.width)
                    colElem.width = col.width;
                
                result.appendChild(colElem);
            }

            return result;
        }

        renderTableRow(elem) {
            var result = this.htmlDocument.createElement("tr");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem.style, result);

            return result;
        }

        renderTableCell(elem: IDomTableCell) {
            var result = this.htmlDocument.createElement("td");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem.style, result);

            if (elem.span) result.colSpan = elem.span;

            return result;
        }

        renderStyleValues(style: IDomStyleValues, ouput: HTMLElement) {
            if (style == null)
                return;

            for (var key in style) {
                if (style.hasOwnProperty(key)) {
                    ouput.style[key] = style[key];
                }
            }
        }

        renderClass(input: IDomElement, ouput: HTMLElement) {
            if (input.className)
                ouput.className = input.className;
        }
    }
}