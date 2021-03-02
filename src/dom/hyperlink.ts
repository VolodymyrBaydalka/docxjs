import globalXmlParser from "../parser/xml-parser";
import { element, fromAttribute } from "../parser/xml-serialize";
import { DocxContainer } from "./dom";

@element('hyperlink')
export class HyperlinkElement extends DocxContainer {
    @fromAttribute('anchor')
    anchor?: string;

    protected parse(elem: Element) {
        this.anchor = globalXmlParser.attr(elem, "anchor");
    }
}