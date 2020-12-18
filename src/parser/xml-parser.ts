import { Length,  LengthUsage, LengthUsageType, convertLength  } from "../dom/common";

export class XmlParser {
    parse(xmlString: string, skipDeclaration: boolean = true): Element {
        if (skipDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");

        return <Element>new DOMParser().parseFromString(xmlString, "application/xml").firstChild;
    }

    elements(elem: Element, localName: string = null): Element[] {
        const result = [];

        for (let i = 0, l = elem.childNodes.length; i < l; i++) {
            let c = elem.childNodes.item(i);

            if (c.nodeType == 1 && (localName == null || (c as Element).localName == localName))
                result.push(c);
        }

        return result;
    }

    element(elem: Element, localName: string): Element {
        for (let i = 0, l = elem.childNodes.length; i < l; i++) {
            let c = elem.childNodes.item(i);

            if (c.nodeType == 1 && (c as Element).localName == localName)
                return c as Element;
        }

        return null;
    }

    attr(elem: Element, localName: string): string {
        for (let i = 0, l = elem.attributes.length; i < l; i++) {
            let a = elem.attributes.item(i);

            if (a.localName == localName)
                return a.value;
        }

        return null;
    }

    intAttr(node: Element, attrName: string, defaultValue: number = null): number {
        var val = this.attr(node, attrName);
        return val ? parseInt(val) : defaultValue;
    }

    floatAttr(node: Element, attrName: string, defaultValue: number = null): number {
        var val = this.attr(node, attrName);
        return val ? parseFloat(val) : defaultValue;
    }

    boolAttr(node: Element, attrName: string, defaultValue: boolean = null) {
        var v = this.attr(node, attrName);

        switch (v) {
            case "1": return true;
            case "0": return false;
            default: return defaultValue;
        }
    }

    lengthAttr(node: Element, attrName: string, usage: LengthUsageType = LengthUsage.Dxa): Length {
        return convertLength(this.attr(node, attrName), usage);
    }
}

const globalXmlParser = new XmlParser();

export default globalXmlParser;