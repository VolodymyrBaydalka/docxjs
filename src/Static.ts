namespace docx {

    export interface Options {
        inWrapper: boolean;
        ignoreWidth: boolean;
        ignoreHeight: boolean;
        debug: boolean;
        className: string;
    }

    export function renderAsync(data, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options = null): PromiseLike<any> {

        var parser = new docx.DocumentParser();
        var renderer = new docx.HtmlRenderer(window.document);

        if (options) {
            parser.ignoreWidth = options.ignoreWidth || parser.ignoreWidth;
            parser.ignoreHeight = options.ignoreHeight || parser.ignoreHeight;
            parser.debug = options.debug || parser.debug;

            renderer.className = options.className || "docx";
        }

        return new JSZip().loadAsync(data)
            .then(zip => {
                var files = [parser.parseDocumentAsync(zip), parser.parseStylesAsync(zip)];
                var num = parser.parseNumberingAsync(zip);
                var rels = parser.parseDocumentRelationsAsync(zip);

                files.push(num || Promise.resolve());
                files.push(rels || Promise.resolve());

                return Promise.all(files);
            })
            .then(parts => {
                var inWrapper = options && options.inWrapper != null ? options.inWrapper : true;
                styleContainer = styleContainer || bodyContainer;

                clearElement(styleContainer);
                clearElement(bodyContainer);

                styleContainer.appendChild(document.createComment("docxjs library predefined styles"));
                styleContainer.appendChild(renderer.renderDefaultStyle());
                styleContainer.appendChild(document.createComment("docx document styles"));
                styleContainer.appendChild(renderer.renderStyles(parts[1]));

                if (parts[2])
                {
                    styleContainer.appendChild(document.createComment("docx document numbering styles"));
                    styleContainer.appendChild(renderer.renderNumbering(parts[2]));
                }

                var documentElement = renderer.renderDocument(parts[0]);

                if (inWrapper) {
                    var wrapper = renderer.renderWrapper();
                    wrapper.appendChild(documentElement);
                    bodyContainer.appendChild(wrapper);
                }
                else {
                    bodyContainer.appendChild(documentElement);
                }

                return { document: parts[0], styles: parts[1], numbering: parts[2], rels: parts[3] };
            });
    }

    function clearElement(elem: HTMLElement) {
        while (elem.firstChild) {
            elem.removeChild(elem.firstChild);
        }
    }
}
