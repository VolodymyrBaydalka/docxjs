import { Length, LengthType, ns, Border, Borders } from "../dom/common";

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

    if (val == null)
        return defaultValue;

    return val === "true" || val === "1";
}

export type LengthUsageType = { mul: number, unit: LengthType };

export const LengthUsage: Record<string, LengthUsageType> = {
    Dxa: { mul: 0.05, unit: "pt" }, //twips
    Emu: { mul: 1 / 12700, unit: "pt" },
    FontSize: { mul: 0.5, unit: "pt" },
    Border: { mul: 0.125, unit: "pt" },
    Percent: { mul: 0.02, unit: "%" },
    LineHeight: { mul: 1 / 240, unit: null }
}

export function lengthAttr(elem: Element, namespaceURI: string, name: string, usage: LengthUsageType = LengthUsage.Dxa): Length {
    var val = elem.getAttributeNS(namespaceURI, name);
    return val ? { value: parseInt(val) * usage.mul, type: usage.unit } : null;
}

export function parseBorder(elem: Element): Border {
    return {
        type: stringAttr(elem, ns.wordml, "val"),
        color: colorAttr(elem, ns.wordml, "color"),
        size: lengthAttr(elem, ns.wordml, "sz", LengthUsage.Border)
    };
}

export function parseBorders(elem: Element): Borders {
    var result = <Borders>{};

    for (let e of elements(elem, ns.wordml)) {
        switch (e.localName) {
            case "left": result.left = parseBorder(e); break;
            case "top": result.top = parseBorder(e); break;
            case "right": result.right = parseBorder(e); break;
            case "botton": result.botton = parseBorder(e); break;
        }
    }

    return result;
}