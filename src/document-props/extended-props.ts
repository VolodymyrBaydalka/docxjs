import { XmlParser } from "../parser/xml-parser";

export interface ExtendedPropsDeclaration {
    template: string,
    totalTime: number,
    pages: number,
    words: number,
    characters: number,
    application: string,
    lines: number,
    paragraphs: number,
    company: string,
    appVersion: string
}

export function parseExtendedProps(root: Element, xmlParser: XmlParser): ExtendedPropsDeclaration {
    const result = <ExtendedPropsDeclaration>{

    };

    for (let el of xmlParser.elements(root)) {
        switch (el.localName) {
            case "Template":
                result.template = el.textContent;
                break;
            case "Pages":
                result.pages = safeParseToInt(el.textContent);
                break;
            case "Words":
                result.words = safeParseToInt(el.textContent);
                break;
            case "Characters":
                result.characters = safeParseToInt(el.textContent);
                break;
            case "Application":
                result.application = el.textContent;
                break;
            case "Lines":
                result.lines = safeParseToInt(el.textContent);
                break;
            case "Paragraphs":
                result.paragraphs = safeParseToInt(el.textContent);
                break;
            case "Company":
                result.company = el.textContent;
                break;
            case "AppVersion":
                result.appVersion = el.textContent;
                break;
        }
    }

    return result;
}

function safeParseToInt(value: string): number {
    if (typeof value === 'undefined')
        return;
    return parseInt(value);
}