import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';

export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
}

export function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null) {
    var parser = new DocumentParser();
    var renderer = new HtmlRenderer(window.document);

    var options = { 
        ignoreHeight: false,
        ignoreWidth: false,
        ignoreFonts: false,
        breakPages: true,
        debug: false,
        experimental: false,
        className: "docx",
        inWrapper: true,
        ... userOptions
    };

    parser.ignoreWidth = options.ignoreWidth;
    parser.debug = options.debug || parser.debug;

    renderer.className = options.className || "docx";
    renderer.inWrapper = options.inWrapper;

    return WordDocument.load(data, parser).then(doc => {
        renderer.render(doc, bodyContainer, styleContainer, options);
        return doc;
    })
}