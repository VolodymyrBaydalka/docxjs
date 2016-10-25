declare class JSZip {
    static loadAsync(data): PromiseLike<any>;
}

declare class Promise {
    static all(ps);
}

namespace docx {
    export function renderAsync(data, bodyContainer: HTMLElement, styleContainer: HTMLElement = null): PromiseLike<any> {

        var parser = new docx.DocumentParser();
        var renderer = new docx.HtmlRenderer(window.document);

        return JSZip.loadAsync(data)
            .then(zip => {
                var files = [parser.parseDocumentAsync(zip), parser.parseStylesAsync(zip)];
                var num = parser.parseNumberingAsync(zip);

                if(num) files.push(num);

                return Promise.all(files);
            })
            .then(parts => {
                styleContainer = styleContainer || bodyContainer;

                clearElement(styleContainer);
                clearElement(bodyContainer);

                styleContainer.appendChild(renderer.renderStyles(parts[1]));
                
                if(parts[2])
                    styleContainer.appendChild(renderer.renderStyles(parts[2]));

                bodyContainer.appendChild(renderer.renderDocument(parts[0]));

                return { document: parts[0], styles: parts[1], numbering: parts[2] };
            });
    }

    function clearElement(elem: HTMLElement) {
        while (elem.firstChild) {
            elem.removeChild(elem.firstChild);
        }
    }
}
