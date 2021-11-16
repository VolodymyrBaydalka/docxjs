import { children, element, fromAttribute } from "../parser/xml-serialize";
import { DocxContainer } from "./dom";

@element('Choice')
export class McChoice extends DocxContainer {
    @fromAttribute('Requires')
    requires: string;
}

@element('Fallback')
export class Fallback extends DocxContainer {
}

@element('AlternateContent')
@children(McChoice, Fallback)
export class McAlternateContent extends DocxContainer {
}