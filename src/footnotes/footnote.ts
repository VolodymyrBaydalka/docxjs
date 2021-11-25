import { DocxContainer } from "../document/dom";
import { element, fromAttribute } from "../parser/xml-serialize";

@element("footnote")
export class WmlFootnote extends DocxContainer {
    @fromAttribute("id")
    id: string;
    @fromAttribute("type")
    type: string;
}