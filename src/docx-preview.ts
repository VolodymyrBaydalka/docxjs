import { Document } from './document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';

export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    breakPages: boolean;
    debug: boolean;
    className: string;
}

export function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null) {
    var parser = new DocumentParser();
    var renderer = new HtmlRenderer(window.document);

    var options = { 
        ignoreHeight: false,
        ignoreWidth: false,
        breakPages: true,
        debug: false,
        className: "docx",
        inWrapper: true,
        ... userOptions
    };

    parser.ignoreWidth = options.ignoreWidth;
    parser.debug = options.debug || parser.debug;

    renderer.className = options.className || "docx";
    renderer.inWrapper = options.inWrapper;

    return Document.load(data, parser).then(doc => {
        renderer.render(doc, bodyContainer, styleContainer, options);
        return doc;
    })
}