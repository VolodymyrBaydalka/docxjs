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

export abstract class StyleWriterBase {
    abstract write(prop: string, value: string): StyleWriterBase;
}

export class TextStyleWriter extends StyleWriterBase {
    constructor(public text: string = "") {
        super();
    }

    write(prop: string, value: string) {
        this.text += `${prop}: ${value};`;
        return this;
    }
}

export class ElementStyleWriter extends StyleWriterBase {
    constructor(public element: HTMLElement) {
        super();
    }

    write(prop: string, value: string) {
        this.element.style[prop] = value;
        return this;
    }
}