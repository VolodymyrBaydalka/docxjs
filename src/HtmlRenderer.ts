module docx {
    export class HtmlRenderer {
        constructor(public docxDocument: IDomDocument, public htmlDocument: HTMLDocument) {

        }

        renderBody(into?: HTMLElement): HTMLElement {
            var bodyElement = document.createElement("section");
            this.renderChildren(this.docxDocument, bodyElement);

            this.renderStyleValues(this.docxDocument, bodyElement.style);

            if(into)
                into.appendChild(bodyElement);

            return bodyElement;
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
            this.renderStyleValues(elem.style, result.style);

            return result;
        }

        renderRun(elem: IDomRun) {

            if (elem.isBreak)
                return this.htmlDocument.createElement("br");

            var result = this.htmlDocument.createElement("span");

            this.renderStyleValues(elem.style, result.style);

            result.textContent = elem.text;

            return result;
        }

        renderTable(elem) {
            var result = this.htmlDocument.createElement("table");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem.style, result.style);

            return result;
        }

        renderTableRow(elem) {
            var result = this.htmlDocument.createElement("tr");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem.style, result.style);

            return result;
        }

        renderTableCell(elem) {
            var result = this.htmlDocument.createElement("td");

            this.renderChildren(elem, result);
            this.renderStyleValues(elem.style, result.style);

            return result;
        }

        renderStyleValues(input: { [name: string]: any }, ouput: CSSStyleDeclaration) {
            if (input == null)
                return;

            for (var key in input) {
                if (input.hasOwnProperty(key)) {
                    ouput[key] = input[key];
                }
            }
        }
    }
}