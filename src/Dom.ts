
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
        style?: IDomStyleValues;
    }

    export interface IDomRun extends IDomElement {
        isBreak?: boolean;
        text?: string;
    }

    export interface IDomDocument extends IDomElement {
    }

    export interface IDomStyle {
        id: string;
        target: string;
        basedOn?: string;
        isDefault?: boolean;
    }

    export interface IDomStyleValues {
        [name: string]: string;
    }
}
