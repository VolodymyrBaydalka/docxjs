
module docx {

    export enum DomType {
        Document,
        Paragraph,
        Run,
        Break,
        Table,
        Row,
        Cell
    }

    export interface IDomElement {
        domType: DomType;
        children?: IDomElement[];
        style?: { [name: string]: any };
    }

    export interface IDomRun extends IDomElement {
        isBreak?: boolean;
        text?: string;
    }

    export interface IDomDocument extends IDomElement {
        styles: { [id: string]: any };
    }
}
