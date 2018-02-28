
namespace docx {

    export enum DomType {
        Document,
        Paragraph,
        Run,
        Break,
        Table,
        Row,
        Cell,
        Hyperlink,
        Drawing,
        Image
    }

    export enum DomRelationshipType {
        Settings,
        Theme,
        StylesWithEffects,
        Styles,
        FontTable,
        Image,
        WebSettings,
        Unknown
    }

    export interface IDomRelationship {
        id: string;
        type: DomRelationshipType;
        target: string;
    }

    export interface IDomElement {
        domType: DomType;
        children?: IDomElement[];
        style?: IDomStyleValues;
        className?: string;
    }

    export interface IDomParagraph extends IDomElement {
        numberingId?: string;
        numberingLevel?: number;
    }

    export interface IDomHyperlink extends IDomElement {
        href?: string;
    }

    export interface IDomRun extends IDomElement {
        id?: string; 
        break?: string;
        wrapper?: string;
        text?: string;
        href?: string;
    }

    export interface IDomTable extends IDomElement {
        columns?: IDomTableColumn[];
        cellStyle?: IDomStyleValues;
    }

    export interface IDomTableRow extends IDomElement {
    }

    export interface IDomTableCell extends IDomElement {
        span?: number;
    }

    export interface IDomDocument extends IDomElement {
    }

    export interface IDomImage extends IDomDocument {
        src: string;
    }

    export interface IDomTableColumn {
        width?: string;
    }

    export interface IDomStyle {
        id: string;
        name?: string;
        target: string;
        basedOn?: string;
        isDefault?: boolean;
        styles: IDomSubStyle[];
    }

    export interface IDomSubStyle {
        target: string;
        values: IDomStyleValues;
    }

    export interface IDomNumbering {
	    id: string;
        level: number;
        style: IDomStyleValues; 
        levelText?: string;
        format?: string;
    }

    export interface IDomStyleValues {
        [name: string]: string;
    }
}
