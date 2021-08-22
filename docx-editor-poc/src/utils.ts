import { DocxElement } from "../../src/document/dom";

export function preventAndstop(e: Event) {
    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();
    return false;
}

export function getSelectionRanges(): Range[] {
    const selection = window.getSelection();
    const result = [];

    for (let i = 0, l = selection.rangeCount; i < l; i++) {
        result.push(selection.getRangeAt(i));
    }

    return result;
}

export function getDocxElement(elem: Node): DocxElement {
    return (elem as any).$$docxElement;
}

export function getXmlElement(elem: DocxElement): Element {
    return (elem as any).$$xmlElement;
}

export function setXmlElement(elem: DocxElement, xml: Node) {
    (elem as any).$$xmlElement = xml;
}