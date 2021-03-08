import { element, fromText } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element('t')
export class TextElement extends DocxElement {
    @fromText()
    text: string;
}