import { XmlParser } from "../parser/xml-parser";
import { children, element } from "../parser/xml-serialize";
import { Border, parseBorder } from "./border";
import { BreakElement } from "./break";
import { Length, LengthUsage, Underline } from "./common";
import { DocxContainer } from "./dom";
import { FieldCharElement } from "./fieldChar";
import { InstructionTextElement } from "./instructions";
import { SymbolElement } from "./symbol";
import { TabElement } from "./tab";
import { TextElement } from "./text";

@element('r')
@children(TextElement, SymbolElement, TabElement, BreakElement, InstructionTextElement, FieldCharElement)
export class RunElement extends DocxContainer {
    id?: string;
    styleName: string;
    props: RunProperties = <RunProperties>{};
}

export interface RunProperties {
    styleName: string;
    fontSize: Length;
    color: string;
    bold: boolean;
    italics: boolean;
    caps: boolean;
    smallCaps: boolean;
    strike: boolean;
    doubleStrike: boolean;
    outline: boolean;
    imprint: boolean;
    underline: Underline;
    border: Border;
    fonts: RunFonts;
    shading: Shading;
    highlight: string;
    spacing: Length;
    stretch: number;
    verticalAlignment: 'baseline' | 'superscript' | 'subscript' | string;
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
        case 'rStyle': 
            props.styleName = xml.attr(elem, 'val');
            break;

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

        case 'smallCaps':
            props.smallCaps = xml.boolAttr(elem, "val", true);
            break;

        case 'imprint':
            props.imprint = xml.boolAttr(elem, "val", true);
            break;

        case 'outline':
            props.outline = xml.boolAttr(elem, "val", true);
            break;

        case 'vertAlign':
            props.verticalAlignment = xml.attr(elem, 'val');
            break;
        
        case 'emboss':
        case 'shadow':
        case 'vanish':
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