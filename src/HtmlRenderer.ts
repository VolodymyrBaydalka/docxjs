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
            for(let style of styles){
                style.id = this.processClassName(style.id);
                style.basedOn = this.processClassName(style.basedOn);
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

                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "border-left");
                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "border-right");
                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "border-top");
                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "border-bottom");

                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "padding-left");
                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "padding-right");
                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "padding-top");
                    c.style = this.copyStyleProperty(table.cellStyle, c.style, "padding-bottom");

                    this.processElement(c);
                }
            }
        }

        copyStyleProperty(input: IDomStyleValues, output: IDomStyleValues, attr: string): IDomStyleValues {
            if (input && input[attr] != null && (output == null || output[attr] == null)) {

                if (output == null)
                    output = {};

                output[attr] = input[attr];
            }

            return output;;
        }

        renderDocument(document: IDomDocument): HTMLElement {
            var bodyElement = this.htmlDocument.createElement("section");

            this.processElement(document);
            this.renderChildren(document, bodyElement);

            this.renderStyleValues(document, bodyElement);

            return bodyElement;
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

        renderParagraph(elem) {
            var result = this.htmlDocument.createElement("p");

            this.renderClass(elem, result);
            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);

            return result;
        }

        renderRun(elem: IDomRun) {

            if (elem.break)
                return this.htmlDocument.createElement(elem.break == "page" ? "hr" : "br");

            var result = this.htmlDocument.createElement("span");

            this.renderStyleValues(elem, result);

            result.textContent = elem.text;

            return result;
        }

        renderTable(elem) {
            var result = this.htmlDocument.createElement("table");

            this.renderClass(elem, result);
            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);

            return result;
        }

        renderTableRow(elem) {
            var result = this.htmlDocument.createElement("tr");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);

            return result;
        }

        renderTableCell(elem: IDomTableCell) {
            var result = this.htmlDocument.createElement("td");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);

            if (elem.span) result.colSpan = elem.span;
            if (elem.vAlign) result.vAlign = elem.vAlign;

            return result;
        }

        renderStyleValues(input: IDomElement, ouput: HTMLElement) {
            if (input.style == null)
                return;

            for (var key in input.style) {
                if (input.style.hasOwnProperty(key)) {
                    ouput.style[key] = input.style[key];
                }
            }
        }

        renderClass(input: IDomElement, ouput: HTMLElement) {
            if (input.className)
                ouput.className = input.className;
        }
    }
}