import { Length } from "../dom/common";

export function forEachElementNS(elem: Element, namespaceURI: string, callback: (elem: Element) => any) {
    elem.childNodes.forEach(n => {
        if(n.nodeType == 1 && n.namespaceURI == namespaceURI)
            callback(<Element>n);
    });
}

export function getAttributeIntValue(elem: Element, namespaceURI: string, name: string): number {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? parseInt(val) : null;
}

export function getAttributeBoolValue(elem: Element, namespaceURI: string, name: string, defaultValue: boolean = false): boolean {
    var val = elem.getAttributeNS(namespaceURI, name);

    if(val == null)
        return defaultValue;

    return val === "true" || val === "1";
}

export function getAttributeLengthValue(elem: Element, namespaceURI: string, name: string, usage: LengthUsage = LengthUsage.Dxa): Length {
    return parseLength(elem.getAttributeNS(namespaceURI, name), usage);
}

export enum LengthUsage {
    Dxa,
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