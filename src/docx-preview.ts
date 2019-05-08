import { Document } from './document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';

export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    debug: boolean;
    className: string;
}

export function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Partial<Options> = null): PromiseLike<any> {
    var parser = new DocumentParser();
    var renderer = new HtmlRenderer(window.document);

    if (options) {
        parser.ignoreWidth = options.ignoreWidth || parser.ignoreWidth;
        parser.ignoreHeight = options.ignoreHeight || parser.ignoreHeight;
        parser.debug = options.debug || parser.debug;

        renderer.className = options.className || "docx";
        renderer.inWrapper = options.inWrapper != null ? options.inWrapper : true;
    }

    return Document.load(data, parser)
        .then(doc => {
            renderer.render(doc, bodyContainer, styleContainer);
            return doc;
        });
}