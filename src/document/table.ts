import globalXmlParser, { attr, elements, XmlParser } from "../parser/xml-parser";
import { element, fromElement } from "../parser/xml-serialize";
import { Length } from "./common";
import { DocxContainer } from "./dom";

@element("tbl")
export class WmlTable extends DocxContainer {
    @fromElement("tblGrid", parseTableColumns)
    columns?: TableColumn[];
    @fromElement("tblPr", parseTableProperties)
    props: TableProperties;

    cellStyle?: Record<string, string>;
}

export interface TableColumn {
    width?: Length;
}

export interface TableProperties {
    alignment: string;
    caption: string;
    tableLook: TableLook;
}

export function parseTableProperties(elem: Element): TableProperties {
    const result = {} as TableProperties;

    for (const e of elements(elem)) {
        switch(e.localName) {
            case "jc":
                result.alignment = attr(e, "val");
                break;
                
            case "tblCaption":
                result.caption = attr(e, "val");
                break;

            case "tblLook":
                result.tableLook = parseTableLook(e);
                break;
        }
    }

    return result;
}

export interface TableLook {
    firstColumn: boolean;
    firstRow: boolean;
    lastColumn: boolean;
    lastRow: boolean;
    noHBand: boolean;
    noVBand: boolean;
}

export function parseTableLook(elem: Element, xml: XmlParser = globalXmlParser): TableLook {
    //TODO
    const intVal = xml.intAttr(elem, "val");

    return {
        firstColumn: xml.boolAttr(elem, 'firstColumn'),
        firstRow: xml.boolAttr(elem, 'firstRow'),
        lastColumn: xml.boolAttr(elem, 'lastColumn'),
        lastRow: xml.boolAttr(elem, 'lastRow'),
        noHBand: xml.boolAttr(elem, 'noHBand'),
        noVBand: xml.boolAttr(elem, 'noVBand')
    }
}

export function parseTableColumns(elem: Element, xml: XmlParser = globalXmlParser): TableColumn[] {
    return xml.elements(elem, 'gridCol').map(e => (<TableColumn>{
        width: xml.lengthAttr(e, "w")
    }));
}