import { XmlParser } from "../parser/xml-parser";

export const ns = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    drawingml: "http://schemas.openxmlformats.org/drawingml/2006/main",
    picture: "http://schemas.openxmlformats.org/drawingml/2006/picture"
}

export type LengthType = "px" | "pt" | "%";

export interface Length {
    value: number;
    type: LengthType
}

export interface Font {
    name: string;
    family: string;
}

export interface CommonProperties {
    fontSize: Length;
    color: string;
}

export type LengthUsageType = { mul: number, unit: LengthType };

export const LengthUsage: Record<string, LengthUsageType> = {
    Dxa: { mul: 0.05, unit: "pt" }, //twips
    Emu: { mul: 1 / 12700, unit: "pt" },
    FontSize: { mul: 0.5, unit: "pt" },
    Border: { mul: 0.125, unit: "pt" },
    Point: { mul: 1, unit: "pt" },
    Percent: { mul: 0.02, unit: "%" },
    LineHeight: { mul: 1 / 240, unit: null }
}

export function convertLength(val: string, usage: LengthUsageType = LengthUsage.Dxa): Length {
    if (!val) {
        return null;
    }

    //"simplified" docx documents use pt's as units
    if (val.endsWith('pt')) {
        return { value: parseFloat(val), type: 'pt' };
    }

    if (val.endsWith('%')) {
        return { value: parseFloat(val), type: '%' };
    }

    return { value: parseInt(val) * usage.mul, type: usage.unit };
}

export function convertBoolean(v: string, defaultValue = false): boolean {
    switch (v) {
        case "1": return true;
        case "0": return false;
        case "true": return true;
        case "false": return false;
        default: return defaultValue;
    }
}

export function convertPercentage(val: string): number {
    return val ? parseInt(val) / 100 : null;
}

export function parseCommonProperty(elem: Element, props: CommonProperties, xml: XmlParser): boolean {
    if(elem.namespaceURI != ns.wordml)
        return false;

    switch(elem.localName) {
        case "color": 
            props.color = xml.attr(elem, "val");
            break;

        case "sz":
            props.fontSize = xml.lengthAttr(elem, "val", LengthUsage.FontSize);
            break;

        default:
            return false;
    }

    return true;
}