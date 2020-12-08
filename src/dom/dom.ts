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
    Symbol = "symbol",
    BookmarkStart = "bookmarkStart",
    BookmarkEnd = "bookmarkEnd"
}

export interface OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[];
    cssStyle?: Record<string, string>;
    className?: string;
    parent?: OpenXmlElement;
}

export interface IDomHyperlink extends OpenXmlElement {
    href?: string;
}


export interface BreakElement extends OpenXmlElement{
    break: "page" | "lastRenderedPageBreak" | "textWrapping";
}

export interface TextElement extends OpenXmlElement{
    text: string;
}

export interface SymbolElement extends OpenXmlElement {
    font: string;
    char: string;
}

export interface IDomTable extends OpenXmlElement {
    columns?: IDomTableColumn[];
    cellStyle?: Record<string, string>;
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
