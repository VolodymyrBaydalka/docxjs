module docx {
    export class HtmlRenderer {
        constructor(public htmlDocument: HTMLDocument) {
        }

        renderDocument(document: IDomDocument): HTMLElement {
            var bodyElement = this.htmlDocument.createElement("section");
            this.renderChildren(document, bodyElement);

            this.renderStyleValues(document, bodyElement);

            return bodyElement;
        }

        renderStyles(styles: IDomStyle[]): HTMLElement {
            var styleElement = document.createElement("style");
            var styleText = "";

            styleElement.type = "text/css";

            for (let style of styles) {

                if (style.isDefault)
                    styleText += style.target + ", ";

                styleText += style.target + "." + style.id + "{\r\n";

                styleText += "}\r\n";
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

            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);

            return result;
        }

        renderRun(elem: IDomRun) {

            if (elem.isBreak)
                return this.htmlDocument.createElement("br");

            var result = this.htmlDocument.createElement("span");

            this.renderStyleValues(elem, result);

            result.textContent = elem.text;

            return result;
        }

        renderTable(elem) {
            var result = this.htmlDocument.createElement("table");

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

        renderTableCell(elem) {
            var result = this.htmlDocument.createElement("td");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem, result);

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
    }
}