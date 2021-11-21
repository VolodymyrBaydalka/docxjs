import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';

export interface Options {
    trimXmlDeclaration: boolean;
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    ignoreLastRenderedPageBreak: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    keepOrigin: boolean;
}

export const defaults = {
    trimXmlDeclaration: false,
    keepOrigin: false,
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    ignoreLastRenderedPageBreak: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
}

export function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null) {
    var parser = new DocumentParser();
    var renderer = new HtmlRenderer(window.document);

    var options: Options = { 
        ...defaults,
        ...userOptions
    };

    Object.assign(parser, options);
    Object.assign(renderer, options);

    return WordDocument.load(data, parser, options).then(doc => {
        renderer.render(doc, bodyContainer, styleContainer, options);
        return doc;
    })
}