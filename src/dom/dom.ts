export class DocxElement {
    className: string = null;
    cssStyle: Record<string, string> = {};

    constructor(public parent?: DocxElement) {
    }
}

export class DocxContainer extends DocxElement {
    children: DocxElement[] = [];
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
