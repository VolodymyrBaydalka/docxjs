import { XmlParser } from "../parser/xml-parser";

export interface CorePropsDeclaration {
    title: string,
    description: string,
    subject: string,
    creator: string,
    keywords: string,
    language: string,
    lastModifiedBy: string,
    revision: number,
}

export function parseCoreProps(root: Element, xmlParser: XmlParser): CorePropsDeclaration {
    const result = <CorePropsDeclaration>{};

    for (let el of xmlParser.elements(root)) {
        switch (el.localName) {
            case "title": result.title = el.textContent; break;
            case "description": result.description = el.textContent; break;
            case "subject": result.subject = el.textContent; break;
            case "creator": result.creator = el.textContent; break;
            case "keywords": result.keywords = el.textContent; break;
            case "language": result.language = el.textContent; break;
            case "lastModifiedBy": result.lastModifiedBy = el.textContent; break;
            case "revision": el.textContent && (result.revision = parseInt(el.textContent)); break;
        }
    }

    return result;
}