
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
        className?: string;
    }

    export interface IDomRun extends IDomElement {
        isBreak?: boolean;
        text?: string;
    }

    export interface IDomTableCell extends IDomElement {
        span?: number;
        vAlign?: string;
    }

    export interface IDomDocument extends IDomElement {
    }

    export interface IDomStyle {
        id: string;
        target: string;
        basedOn?: string;
        isDefault?: boolean;
        styles: IDomSubStyle[];
    }

    export interface IDomSubStyle {
        target: string;
        values: IDomStyleValues;
    }

    export interface IDomStyleValues {
        [name: string]: string;
    }
}
