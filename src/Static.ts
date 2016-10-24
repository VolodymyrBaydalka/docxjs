declare class JSZip {
    static loadAsync(data): PromiseLike<any>;
}

module docx {
    export function parseAsync(data) {
        var parser = new DocumentParser();

        return JSZip.loadAsync(data).then(zip => parser.parseAsync(zip));
    }

    export function renderAsync(data, bodyContainer: HTMLElement, styleContainer?: HTMLElement, document?: HTMLDocument) {
        return parseAsync(data).then(x => {
            var renderer = new HtmlRenderer(x, document || window.document);

            renderer.renderBody(bodyContainer);

            return x;
        });
    }
}
