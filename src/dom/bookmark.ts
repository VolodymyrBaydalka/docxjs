import { element, fromAttribute } from "../parser/xml-serialize";
import { DocxElement } from "./dom";

@element("bookmarkStart")
export class BookmarkStartElement extends DocxElement {
    @fromAttribute("id")
    id: string;
    @fromAttribute("name")
    name: string;
    @fromAttribute("colFirst")
    colFirst: number;
    @fromAttribute("colLast")
    colLast: number;
}

@element("bookmarkEnd")
export class BookmarkEndElement extends DocxElement {
    @fromAttribute("id")
    id: string;
}