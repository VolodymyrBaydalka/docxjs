
module docx {
    export class DocumentParser {
        public skipDeclaration: boolean = true;

        parseAsync(zip): IDomDocument {
            var result: IDomDocument = {
                domType: DomType.Document,
                children: [],
                styles: {}
            };

            var documentPromise = zip.files["word/document.xml"]
                .async("string")
                .then((xml) => this.parseDocumentFile(xml));

            return documentPromise;
        }

        private parseDocumentFile(xmlString) {
            var result: IDomDocument = {
                domType: DomType.Document,
                children: [],
                style: {},
                styles: {}
            };

            var xbody = findXmlNode(parseXML(xmlString, this.skipDeclaration), "body");

            if (xbody.firstChild.localName == "parsererror")
                console.log(xbody.innerText);

            for (var i = 0; i < xbody.childNodes.length; i++) {
                var node = xbody.childNodes[i];

                switch (node.localName) {
                    case "p":
                        result.children.push(parseParagraph(node));
                        break;

                    case "tbl":
                        result.children.push(parseTable(node));
                        break;

                    case "sectPr":
                        parseSectionProperties(node, result);                        
                        break;
                }
            }

            return result;
        }
    }

    function parseSectionProperties(node: Node, elem:IDomElement){
        forEachNodes(node, n => {
            switch(n.localName){
                case "pgMar":
                    elem.style["marginLeft"] = parseInt(xmlAttrValue(n, "left")) / 20 + "pt";
                    elem.style["marginRight"] = parseInt(xmlAttrValue(n, "right")) / 20 + "pt";
                    elem.style["marginTop"] = parseInt(xmlAttrValue(n, "top")) / 20 + "pt";
                    elem.style["marginBottom"] = parseInt(xmlAttrValue(n, "bottom")) / 20 + "pt";
                    break;
            }
        });
    }

    function parseStyles(node: Node) {
    }

    function forEachNodes(node:Node, cb: (n: Node) => void){
        for (var i = 0; i < node.childNodes.length; i++) 
            cb(node.childNodes[i]);
    }

    function parseParagraph(node: Node): IDomElement {
        var result: IDomElement = { domType: DomType.Paragraph, children: [] };

        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];
            switch (c.localName) {
                case "r":
                    result.children.push(parseRun(c));
                    break;

                case "pPr":
                    result.style = parseDefaultProperties(c, {}, null);
                    break;
            }
        }

        return result;
    }

    function parseRun(node: Node): IDomElement {
        var result: IDomRun = { domType: DomType.Run };

        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];

            switch (c.localName) {
                case "t":
                    result.text = c.textContent;
                    break;

                case "br":
                    result.isBreak = true;
                    break;

                case "rPr":
                    result.style = parseDefaultProperties(c, {}, null);
                    break;
            }
        }

        return result;
    }

    function parseTable(node: Node): IDomElement {
        var result: IDomElement = { domType: DomType.Table, children: [] };

        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];

            switch (c.localName) {
                case "tr": result.children.push(parseTableRow(c)); break;
            }
        }

        return result;
    }

    function parseTableRow(node: Node): IDomElement {
        var result: IDomElement = { domType: DomType.Row, children: [] };

        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];

            switch (c.localName) {
                case "tc": result.children.push(parseTableCell(c)); break;
            }
        }

        return result;
    }
    function parseTableCell(node: Node): IDomElement {
        var result: IDomElement = { domType: DomType.Cell, children: [] };

        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];

            switch (c.localName) {
                case "tbl": result.children.push(parseTable(c)); break;
                case "p": result.children.push(parseParagraph(c)); break;
            }
        }

        return result;
    }

    function parseDefaultProperties(node: Node, output: { [name: string]: any }, handler: (prop: Node) => void): { [name: string]: any } {
        for (var i = 0; i < node.childNodes.length; i++) {
            var c = node.childNodes[i];

            switch (c.localName) {
                case "jc":
                    output["text-align"] = xmlAttrValue(c, "val");
                    break;

                case "color":
                    output["color"] = xmlAttrValue(c, "val");
                    break;

                case "b":
                    output["font-weight"] = "bold";
                    break;

                default:
                    if (handler != null)
                        handler(node);
                    break;
            }

        }

        return output;
    }

    function parseXML(xmlString, skipDeclaration) {
        if (skipDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");

        return new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
    }

    function findXmlNode(parent, tagName) {
        for (var i = 0; i < parent.childNodes.length; i++) {
            var element = parent.childNodes.item(i);

            if (element.localName == tagName)
                return element;
        }

        return null;
    }

    function xmlAttrValue(node, attrName) {
        for (var i = 0; i < node.attributes.length; i++) {
            var attr = node.attributes.item(i);

            if (attr.localName == attrName)
                return attr.value;
        }

        return null;
    }

    function twentiethToPoint(val) {
        return val / 20;
    }

    function halfToPoint(val) {
        return val / 2;
    }

    function fiftiethsToPercent(val) {
        return val / 50;
    }
}
