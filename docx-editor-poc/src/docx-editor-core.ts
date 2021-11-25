import { DocumentParser } from "../../src/document-parser";
import { DocxContainer, DocxElement } from "../../src/document/dom";
import { WmlRun } from "../../src/document/run";
import { WmlText } from "../../src/document/text";
import { HtmlRenderer } from "../../src/html-renderer";
import { WordDocument } from "../../src/word-document";
import { splitRuns } from "./docx-utils";
import { getDocxElement, getSelectionRanges, getXmlElement, preventAndstop } from "./utils";


export class DocxEditorCore {

    fileName: string;
    document: WordDocument;
    renderer: HtmlRenderer;
    parser: DocumentParser;

    mutationObserver: MutationObserver;
    modified: boolean = false;

    contentContainer: HTMLElement;
    styleContainer: HTMLElement;

    constructor() {
        this.mutationObserver = new MutationObserver(this.mutationCallback.bind(this)); 

        this.parser = new DocumentParser({ keepOrigin: true });
        this.renderer = new HtmlRenderer(window.document);
    }

    init(contentContainer: HTMLElement, styleContainer: HTMLElement) {
        this.contentContainer = contentContainer;
        this.styleContainer = styleContainer;

        contentContainer.setAttribute("contentEditable", "");
        contentContainer.addEventListener("keydown", this.onKeyDown);
    }

    async open(content: File) {
        this.modified = false;

        this.mutationObserver.disconnect();

        this.fileName = content.name;
        this.document = await WordDocument.load(content, this.parser, {
            keepOrigin: true
        });
        this.renderer.render(this.document, this.contentContainer, this.styleContainer, {
            ignoreHeight: true,
            breakPages: false,
            keepOrigin: true
        } as any);

        this.mutationObserver.observe(this.contentContainer, {
            attributes: false,
            childList: true,
            subtree: true,
            characterData: true
        });
    }

    async save(): Promise<Blob> {

        if (this.modified) {
            this.document.documentPart.save();
        }

        const blob = await this.document.save();
        return this.fileName ? new File([blob], this.fileName) : blob;
    }

    private mutationCallback(mutations: MutationRecord[]) {

        for(let mutation of mutations) {
            const target = mutation.target;
            const docxElement = getDocxElement(target);

            switch(mutation.type) {
                case "characterData":
                    this.updateText(docxElement, target.textContent);
                    break;

                case "childList":
                    const elems = Array.from(mutation.removedNodes).map(n => (n as any).$$docxElement).filter(x => x);
                    this.removeChildren(docxElement as DocxContainer, elems);
                    break;
            }

            console.log(mutation);
        }

        this.modified = true;
    }

    private updateText(elem: DocxElement, text: string) {
        getXmlElement(elem).textContent = text;
    }

    private removeChildren(elem: DocxContainer, children: DocxElement[]) {
        elem.children = elem.children.filter(c => !children.includes(c));

        for(let cs of children.map(c => (c as any).$$source).filter(x => x)) {
            cs.remove();
        }
    }

    private onKeyDown(event: KeyboardEvent) {
        if (event.ctrlKey) {
            return preventAndstop(event);
        }

        return true;
    }

    bold() {
        for(let r of getSelectionRanges()) {
            const text = getDocxElement(r.startContainer) as WmlText;

            splitRuns(text, r.startOffset);
        }
        
        this.modified = true;
    }

    private findOrCreateElement(elem: Element, localName: string): Element {
        let result = Array.from(elem.childNodes).find(e => e.nodeType === 1 && (e as Element).localName == localName);

        if (result == null) {
            result = elem.ownerDocument.createElementNS(elem.namespaceURI, localName);
            elem.insertBefore(result, elem.firstChild);
        }

        return result as Element;
    }
}