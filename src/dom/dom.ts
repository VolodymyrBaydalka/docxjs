export class DocxElement {
    className: string = null;
    cssStyle: Record<string, string> = {};
    parent: DocxElement = null;

    constructor(elem?: Element) {
        this.init();
        
        if (elem) {
            this.parse(elem);
        }
    }

    protected init() {
    }

    protected parse(elem: Element) {
    }
}

export class DocxContainer extends DocxElement {
    children: DocxElement[];

    protected init() {
        this.children = [];
    }
}

export interface IDomNumbering {
    id: string;
    level: number;
    style: Record<string, string>;
    levelText?: string;
    format?: string;
    bullet?: NumberingPicBullet;
}

export interface NumberingPicBullet {
    id: number;
    src: string;
    style?: string;
}
