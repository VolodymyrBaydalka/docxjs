export class XmlParser {
    parse(xmlString: string, skipDeclaration: boolean = true): Element {
        if (skipDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");

        return <Element>new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
    }

    elements(elem: Element): Element[] {
        const result = [];

        for(let i = 0, l = elem.childNodes.length; i < l; i ++) {
            let c = elem.childNodes.item(i);
            
            if(c.nodeType == 1)
                result.push(c);
        }

        return result;
    }

    element(elem: Element, localName: string): Element {
        for(let i = 0, l = elem.childNodes.length; i < l; i ++) {
            let c = elem.childNodes.item(i);
            
            if(c.nodeType == 1 && c.nodeName == localName)
                return c as Element;
        }

        return null;
    }

    attr(elem: Element, localName: string): string {
        for(let i = 0, l = elem.attributes.length; i < l; i ++) {
            let a = elem.attributes.item(i);
            
            if(a.localName == localName)
                return a.value;
        }

        return null;      
    }
}