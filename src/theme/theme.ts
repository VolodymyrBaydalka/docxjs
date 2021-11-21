import { XmlParser } from "../parser/xml-parser";

export class DmlTheme {
    colorScheme: DmlColorScheme;
    fontScheme: DmlFontScheme;
}

export interface DmlColorScheme {
    name: string;
    colors: Record<string, string>;
}

export interface DmlFontScheme {
    name: string;
    majorFont: DmlFormInfo,
    minorFont: DmlFormInfo
}

export interface DmlFormInfo {
    latinTypeface: string;
    eaTypeface: string;
    csTypeface: string;
}

export function parseTheme(elem: Element, xml: XmlParser) {
    var result = new DmlTheme();
    var themeElements = xml.element(elem, "themeElements");

    for (let el of xml.elements(themeElements)) {
        switch(el.localName) {
            case "clrScheme": result.colorScheme = parseColorScheme(el, xml); break;
            case "fontScheme": result.fontScheme = parseFontScheme(el, xml); break;
        }
    }

    return result;
}

export function parseColorScheme(elem: Element, xml: XmlParser) {
    var result: DmlColorScheme = { 
        name: xml.attr(elem, "name"),
        colors: {}
    };

    for (let el of xml.elements(elem)) {
        var srgbClr = xml.element(el, "srgbClr");
        var sysClr = xml.element(el, "sysClr");

        if (srgbClr) {
            result.colors[el.localName] = xml.attr(srgbClr, "val");
        }
        else if (sysClr) {
            result.colors[el.localName] = xml.attr(sysClr, "lastClr");
        }
    }

    return result;
}

export function parseFontScheme(elem: Element, xml: XmlParser) {
    var result: DmlFontScheme = { 
        name: xml.attr(elem, "name"),
    } as DmlFontScheme;

    for (let el of xml.elements(elem)) {
        switch (el.localName) {
            case "majorFont": result.majorFont = parseFontInfo(el, xml); break;
            case "minorFont": result.minorFont = parseFontInfo(el, xml); break;
        }
    }

    return result;
}

export function parseFontInfo(elem: Element, xml: XmlParser): DmlFormInfo {
    return {
        latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
        eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
        csTypeface: xml.elementAttr(elem, "cs", "typeface"),
    };
}