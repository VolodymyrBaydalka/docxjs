import { DocxElement } from "./dom";

export class TextElement extends DocxElement {
    text: string;

    protected parse(elem: Element) {
        super.parse(elem);
        this.text = elem.textContent;
    }
}