import { Length } from "../dom/common";

export function elements(elem: Element, namespaceURI: string = null, localName: string = null): Element[] {
    let result = [];

    for (let i = 0; i < elem.childNodes.length; i++) {
        let n = elem.childNodes[i];

        if (n.nodeType == 1
            && (namespaceURI == null || n.namespaceURI == namespaceURI)
            && (localName == null || (n as Element).localName == localName))
            result.push(n);
    }

    return result;
}

export function stringAttr(elem: Element, namespaceURI: string, name: string): string {
    return elem.getAttributeNS(namespaceURI, name);
}

export function intAttr(elem: Element, namespaceURI: string, name: string): number {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? parseInt(val) : null;
}

export function colorAttr(elem: Element, namespaceURI: string, name: string): string {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? `#${val}` : null;
}

export function boolAttr(elem: Element, namespaceURI: string, name: string, defaultValue: boolean = false): boolean {
    var val = elem.getAttributeNS(namespaceURI, name);

    if(val == null)
        return defaultValue;

    return val === "true" || val === "1";
}

export function lengthAttr(elem: Element, namespaceURI: string, name: string, usage: LengthUsage = LengthUsage.Dxa): Length {
    return parseLength(elem.getAttributeNS(namespaceURI, name), usage);
}

export enum LengthUsage {
    Dxa, //twips
    Emu,
    FontSize,
    Border,
    Percent
}

export function parseLength(val: string | null, usage: LengthUsage = LengthUsage.Dxa): Length {
    if (!val)
        return null;

    var num = parseInt(val);

    switch (usage) {
        case LengthUsage.Dxa: return { value: 0.05 * num, type: "pt" };
        case LengthUsage.Emu: return { value: num / 12700, type: "pt" };
        case LengthUsage.FontSize: return { value: 0.5 * num, type: "pt" };
        case LengthUsage.Border: return { value: 0.125 * num, type: "pt" };
        case LengthUsage.Percent: return { value: 0.02 * num, type: "%" };
    }

    return null;
}