declare class JSZip {
    static loadAsync(data): PromiseLike<any>;
}

module docx {
    export function renderAsync(data, bodyContainer: HTMLElement, styleContainer: HTMLElement = null): PromiseLike<any> {

        var parser = new docx.DocumentParser();
        var renderer = new docx.HtmlRenderer(window.document);
        var _zip = null;
        var _doc = null;

        return JSZip.loadAsync(data)
            .then(zip => { _zip = zip; return parser.parseDocumentAsync(_zip); })
            .then(doc => { _doc = doc; return parser.parseStylesAsync(_zip); })
            .then(styles => {

                styleContainer = styleContainer || bodyContainer;

                clearElement(styleContainer);
                clearElement(bodyContainer);

                styleContainer.appendChild(renderer.renderStyles(styles));
                bodyContainer.appendChild(renderer.renderDocument(_doc));

                return { document: _doc, styles: styles };
            });
    }

    function clearElement(elem: HTMLElement) {
        while (elem.firstChild) {
            elem.removeChild(elem.firstChild);
        }
    }
}
