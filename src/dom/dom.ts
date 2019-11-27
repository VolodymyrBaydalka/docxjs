export enum DomType {
    Document = "document",
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

export interface IDomTable extends OpenXmlElement {
    cellStyle?: IDomStyleValues;
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
