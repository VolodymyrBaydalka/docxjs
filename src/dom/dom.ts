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

export interface OpenXmlElement {
    domType: DomType;
    children?: OpenXmlElement[];
    style?: IDomStyleValues;
    className?: string;
    parent?: OpenXmlElement;
}

export interface DocxTab {
    style: string;
    leader: string;
    position: string;
}

export interface IDomHyperlink extends OpenXmlElement {
    href?: string;
}

export interface IDomRun extends OpenXmlElement {
    id?: string;
    break?: string;
    wrapper?: string;
    text?: string;
    href?: string;
    tab?: boolean;
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
