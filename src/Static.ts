namespace docx {

    export interface Options {
        inWrapper: boolean;
        ignoreWidth: boolean;
        ignoreHeight: boolean;
        debug: boolean;
    }

    export function renderAsync(data, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options = null): PromiseLike<any> {

        var parser = new docx.DocumentParser();
        var renderer = new docx.HtmlRenderer(window.document);

        if (options) {
            parser.ignoreWidth = options.ignoreWidth || parser.ignoreWidth;
            parser.ignoreHeight = options.ignoreHeight || parser.ignoreHeight;
            parser.debug = options.debug || parser.debug;
        }

        return new JSZip().loadAsync(data)
            .then(zip => {
                var files = [parser.parseDocumentAsync(zip), parser.parseStylesAsync(zip)];
                var num = parser.parseNumberingAsync(zip);

                if (num) files.push(num);

                return Promise.all(files);
            })
            .then(parts => {
                var inWrapper = options && options.inWrapper != null ? options.inWrapper : true;
                styleContainer = styleContainer || bodyContainer;

                clearElement(styleContainer);
                clearElement(bodyContainer);

                styleContainer.appendChild(renderer.renderDefaultStyle());
                styleContainer.appendChild(renderer.renderStyles(parts[1]));

                if (parts[2])
                    styleContainer.appendChild(renderer.renderNumbering(parts[2]));

                var documentElement = renderer.renderDocument(parts[0]);

                if (inWrapper) {
                    var wrapper = renderer.renderWrapper();
                    wrapper.appendChild(documentElement);
                    bodyContainer.appendChild(wrapper);
                }
                else {
                    bodyContainer.appendChild(documentElement);
                }

                return { document: parts[0], styles: parts[1], numbering: parts[2] };
            });
    }

    function clearElement(elem: HTMLElement) {
        while (elem.firstChild) {
            elem.removeChild(elem.firstChild);
        }
    }
}
