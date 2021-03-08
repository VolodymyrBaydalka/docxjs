import globalXmlParser, { XmlParser } from "../parser/xml-parser";
import { element, fromElement } from "../parser/xml-serialize";
import { Length } from "./common";
import { DocxContainer } from "./dom";

@element("tbl")
export class TableElement extends DocxContainer {
    @fromElement("tblGrid", parseTableColumns)
    columns?: TableColumn[];

    cellStyle?: Record<string, string>;
}

export interface TableColumn {
    width?: Length;
}

export function parseTableColumns(elem: Element): TableColumn[] {
    return globalXmlParser.elements(elem, 'gridCol').map(e => (<TableColumn>{
        width: globalXmlParser.lengthAttr(e, "w")
    }));
}