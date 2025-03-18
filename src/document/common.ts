import { XmlParser } from "../parser/xml-parser";
import { clamp } from "../utils";

export const ns = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    drawingml: "http://schemas.openxmlformats.org/drawingml/2006/main",
    picture: "http://schemas.openxmlformats.org/drawingml/2006/picture",
	compatibility: "http://schemas.openxmlformats.org/markup-compatibility/2006",
	math: "http://schemas.openxmlformats.org/officeDocument/2006/math"
}

export type LengthType = "px" | "pt" | "%" | "";
export type Length = string;

export interface Font {
    name: string;
    family: string;
}

export interface CommonProperties {
    fontSize: Length;
    color: string;
}

export type LengthUsageType = { mul: number, unit: LengthType, min?: number, max?: number };

export const LengthUsage: Record<string, LengthUsageType> = {
    Dxa: { mul: 0.05, unit: "pt" }, //twips
    Emu: { mul: 1 / 12700, unit: "pt" },
    FontSize: { mul: 0.5, unit: "pt" },
    Border: { mul: 0.125, unit: "pt", min: 0.25, max: 12 }, //NOTE: http://officeopenxml.com/WPtextBorders.php
    Point: { mul: 1, unit: "pt" },
    Percent: { mul: 0.02, unit: "%" },
    LineHeight: { mul: 1 / 240, unit: "" },
    VmlEmu: { mul: 1 / 12700, unit: "" },
}

export function convertLength(val: string, usage: LengthUsageType = LengthUsage.Dxa): string {
    //"simplified" docx documents use pt's as units
    if (val == null || /.+(p[xt]|[%])$/.test(val)) {
        return val;
    }

    var num = parseInt(val) * usage.mul;

    if (usage.min && usage.max)
        num = clamp(num, usage.min, usage.max);
    
	return `${num.toFixed(2)}${usage.unit}`;
}

export function convertBoolean(v: string, defaultValue = false): boolean {
    switch (v) {
        case "1": return true;
        case "0": return false;
        case "on": return true;
        case "off": return false;
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