import { ElementBase } from "./element-base";
import { fromText, element } from "../parser/xml-serialize";

@element("instrText")
export class InstructionText extends ElementBase {
    @fromText()
    text: string;
}