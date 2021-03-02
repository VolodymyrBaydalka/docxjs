import { element, fromAttribute } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element('fldChar')
export class FieldCharElement extends DocxElement {
    @fromAttribute('fldCharType')
    type: 'begin' | 'end' | 'separate'; 
}