export enum DomType {
    Document = "document",
    Paragraph = "paragraph",
    Run = "run",
    Break = "break",
    Table = "table",
    Row = "row",
    Cell = "cell",
    Hyperlink = "hyperlink",
    Drawing = "drawing",
    Image = "image",
    Text = "text",
    Tab = "tab",
    Symbol = "symbol"
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

export interface OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[];
    style?: IDomStyleValues;
    className?: string;
    parent?: OpenXmlElement;
}

export interface IDomHyperlink extends OpenXmlElement {
    href?: string;
}


export interface TextElement extends OpenXmlElement{
    text: string;
}

export interface SymbolElement extends OpenXmlElement {
    font: string;
    char: string;
}

export interface IDomRun extends OpenXmlElement {
    id?: string;
    break?: string;
    wrapper?: string;
    href?: string;
    fldCharType?: "begin" | "end" | "separate" | string;
    instrText?: string;
}

export interface IDomTable extends OpenXmlElement {
    columns?: IDomTableColumn[];
    cellStyle?: IDomStyleValues;
}

export interface IDomTableRow extends OpenXmlElement {
}

export interface IDomTableCell extends OpenXmlElement {
    span?: number;
}

export interface IDomImage extends OpenXmlElement {
    src: string;
}

export interface IDomTableColumn {
    width?: string;
}

export interface IDomStyle {
    id: string;
    name?: string;
    aliases?: string[];
    target: string;
    basedOn?: string;
    isDefault?: boolean;
    styles: IDomSubStyle[];
    linked?: string;
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
    bullet?: NumberingPicBullet;
}

export interface NumberingPicBullet {
    id: number;
    src: string;
    style?: string;
}

export interface IDomStyleValues {
    [name: string]: string;
}
