import { XmlParser } from "../parser/xml-parser";
import { Border, parseBorder } from "./border";
import { Length, LengthUsage, Underline } from "./common";
import { OpenXmlElement } from "./dom";

export interface RunElement extends OpenXmlElement, RunProperties {
    id?: string;
    break?: string;
    wrapper?: string;
    href?: string;
    fldCharType?: "begin" | "end" | "separate" | string;
    instrText?: string;
}

export interface RunProperties {
    fontSize: Length;
    color: string;
    bold: boolean;
    italics: boolean;
    caps: boolean;
    strike: boolean;
    doubleStrike: boolean;
    underline: Underline;
    border: Border;
    fonts: RunFonts;
    shading: Shading;
    highlight: string;
    spacing: Length;
    stretch: number;
}

export interface Shading {
    foreground: string,
    background: string,
    type: string
}

export interface RunFonts {
    ascii: string;
    hAscii: string;
    cs: string;
    eastAsia: string;
}

export function parseRunProperties(elem: Element, xml: XmlParser): RunProperties {
    let result = <RunProperties>{};

    for(let el of xml.elements(elem)) {
        parseRunProperty(el, result, xml);
    }

    return result;
}

export function parseRunProperty(elem: Element, props: RunProperties, xml: XmlParser) {
    switch (elem.localName) {
        case 'bdr': 
            props.border = parseBorder(elem, xml);
            break;

        case 'rFonts': 
            props.fonts = parseRunFonts(elem, xml);
            break;

        case 'shd': 
            props.shading = parseShading(elem, xml);
            break;

        case 'highlight': 
            props.highlight = xml.attr(elem, 'val');
            break;
        
        case 'spacing':
            props.spacing = xml.lengthAttr(elem, 'val');
            break;

        case 'w':
            props.stretch = xml.percentageAttr(elem, 'val');
            break;

            case "color": 
            props.color = xml.attr(elem, "val");
            break;

        case "sz":
            props.fontSize = xml.lengthAttr(elem, "val", LengthUsage.FontSize);
            break;

        case "b":
            props.bold = xml.boolAttr(elem, "val", true);
            break;

        case "strike":
            props.strike = xml.boolAttr(elem, "val", true);
            break;
    
        case "dstrike":
            props.doubleStrike = xml.boolAttr(elem, "val", true);
            break;
    
        case "i":
            props.italics = xml.boolAttr(elem, "val", true);
            break;

        case "u":
            props.underline = {
                color: xml.attr(elem, "color"),
                type: xml.attr(elem, 'val')
            };
            break;
            
        case 'caps':
            props.caps = xml.boolAttr(elem, "val", true);
            break;

        default:
            return false;
    }

    return true;
}

export function parseRunFonts(elem: Element, xml: XmlParser): RunFonts {
    return {
        ascii: xml.attr(elem, 'ascii'),
        hAscii: xml.attr(elem, 'hAscii'),
        cs: xml.attr(elem, 'cs'),
        eastAsia: xml.attr(elem, 'eastAsia'),
    };
}

export function parseShading(elem: Element, xml: XmlParser): Shading {
    return {
        type: xml.attr(elem, 'val'),
        foreground: xml.attr(elem, 'color'),
        background: xml.attr(elem, 'fill')
    };
}