import { fromText, element } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element("instrText")
export class InstructionTextElement extends DocxElement {
    @fromText()
    text: string;
}