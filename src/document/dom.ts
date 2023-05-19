export enum DomType {
    Document = "document",
    Paragraph = "paragraph",
    Run = "run",
    Break = "break",
    NoBreakHyphen = "noBreakHyphen",
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
    BookmarkEnd = "bookmarkEnd",
    Footer = "footer",
    Header = "header",
    FootnoteReference = "footnoteReference", 
	EndnoteReference = "endnoteReference",
    Footnote = "footnote",
    Endnote = "endnote",
    SimpleField = "simpleField",
    ComplexField = "complexField",
    Instruction = "instruction",
	VmlPicture = "vmlPicture",
	MmlMath = "mmlMath",
	MmlMathParagraph = "mmlMathParagraph",
	MmlFraction = "mmlFraction",
	MmlNumerator = "mmlNumerator",
	MmlDenominator = "mmlDenominator",
	MmlRadical = "mmlRadical",
	MmlBase = "mmlBase",
	MmlDegree = "mmlDegree",
	MmlSuperscript = "mmlSuperscript",
	MmlSubscript = "mmlSubscript",
	MmlSubArgument = "mmlSubArgument",
	MmlSuperArgument = "mmlSuperArgument",
	MmlNary = "mmlNary",
	MmlDelimiter = "mmlDelimiter",
	MmlRun = "mmlRun",
	VmlElement = "vmlElement",
	Inserted = "inserted",
	Deleted = "deleted",
	DeletedText = "deletedText",
	CheckboxFormField = "checkBoxFormField"
}

export interface OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[];
    cssStyle?: Record<string, string>;
    props?: Record<string, any>;
    
	styleName?: string; //style name
	className?: string; //class mods

    parent?: OpenXmlElement;
}

export interface WmlHyperlink extends OpenXmlElement {
	id?: string;
    href?: string;
}

export interface WmlNoteReference extends OpenXmlElement {
    id: string;
}

export interface WmlBreak extends OpenXmlElement{
    break: "page" | "lastRenderedPageBreak" | "textWrapping";
}

export interface WmlText extends OpenXmlElement{
    text: string;
}

export interface WmlSymbol extends OpenXmlElement {
    font: string;
    char: string;
}

export interface WmlTable extends OpenXmlElement {
    columns?: WmlTableColumn[];
    cellStyle?: Record<string, string>;

	colBandSize?: number;
	rowBandSize?: number;
}

export interface WmlTableRow extends OpenXmlElement {
	isHeader?: boolean;
}

export interface WmlTableCell extends OpenXmlElement {
	verticalMerge?: 'restart' | 'continue' | string;
    span?: number;
}

export interface IDomImage extends OpenXmlElement {
    src: string;
}

export interface WmlTableColumn {
    width?: string;
}

export interface IDomNumbering {
    id: string;
    level: number;
    pStyleName: string;
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
