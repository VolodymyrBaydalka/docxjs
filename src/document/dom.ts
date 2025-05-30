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
    SmartTag = "smartTag",
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
	MmlFunction = "mmlFunction",
	MmlFunctionName = "mmlFunctionName",
	MmlNumerator = "mmlNumerator",
	MmlDenominator = "mmlDenominator",
	MmlRadical = "mmlRadical",
	MmlBase = "mmlBase",
	MmlDegree = "mmlDegree",
	MmlSuperscript = "mmlSuperscript",
	MmlSubscript = "mmlSubscript",
	MmlPreSubSuper = "mmlPreSubSuper",
	MmlSubArgument = "mmlSubArgument",
	MmlSuperArgument = "mmlSuperArgument",
	MmlNary = "mmlNary",
	MmlDelimiter = "mmlDelimiter",
	MmlRun = "mmlRun",
	MmlEquationArray = "mmlEquationArray",
	MmlLimit = "mmlLimit",
	MmlLimitLower = "mmlLimitLower",
	MmlMatrix = "mmlMatrix",
	MmlMatrixRow = "mmlMatrixRow",
	MmlBox = "mmlBox",
	MmlBar = "mmlBar",
	MmlGroupChar = "mmlGroupChar",
	VmlElement = "vmlElement",
	Inserted = "inserted",
	Deleted = "deleted",
	DeletedText = "deletedText",
	Comment = "comment",
	CommentReference = "commentReference",
	CommentRangeStart = "commentRangeStart",
	CommentRangeEnd = "commentRangeEnd",
    AltChunk = "altChunk"
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

export abstract class OpenXmlElementBase implements OpenXmlElement {
    type: DomType;
    children?: OpenXmlElement[] = [];
    cssStyle?: Record<string, string> = {};
    props?: Record<string, any>;

    className?: string;
    styleName?: string;

    parent?: OpenXmlElement;
}

export interface WmlHyperlink extends OpenXmlElement {
	id?: string;
    anchor?: string;
}

export interface WmlAltChunk extends OpenXmlElement {
	id?: string;
}

export interface WmlSmartTag extends OpenXmlElement {
	uri?: string;
    element?: string;
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
    gridBefore?: number;
    gridAfter?: number;
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
    start: number;
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
