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
    pStyle: Record<string, string>;
    rStyle: Record<string, string>;
    levelText?: string;
    suff: string;
    format?: string;
    bullet?: NumberingPicBullet;
}

export interface NumberingPicBullet {
    id: number;
    src: string;
    style?: string;
}
