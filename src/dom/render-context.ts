import { ParagraphNumbering } from "./paragraph";

export class RenderContext {
    html: HTMLDocument;
    options: any;
    className: string;
    document: any;

    numberingClass(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }
}