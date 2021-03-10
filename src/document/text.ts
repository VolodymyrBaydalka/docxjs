import { element, fromAttribute, fromText } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element('t')
export class TextElement extends DocxElement {
    @fromText()
    text: string;
}

@element('sym')
export class SymbolElement extends DocxElement {
    @fromAttribute('font')
    font: string;
    @fromAttribute('char')
    char: string;
}

@element('tab')
export class TabElement extends DocxElement {
}

@element("instrText")
export class InstructionTextElement extends DocxElement {
    @fromText()
    text: string;
}