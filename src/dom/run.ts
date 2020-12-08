import { XmlParser } from "../parser/xml-parser";
import { CommonProperties, parseCommonProperty } from "./common";
import { OpenXmlElement } from "./dom";

export interface RunElement extends OpenXmlElement, RunProperties {
    id?: string;
    break?: string;
    wrapper?: string;
    href?: string;
    fldCharType?: "begin" | "end" | "separate" | string;
    instrText?: string;
}

export interface RunProperties extends CommonProperties {

}

export function parseRunProperties(elem: Element, xml: XmlParser): RunProperties {
    let result = <RunProperties>{};

    for(let el of xml.elements(elem)) {
        parseRunProperty(el, result, xml);
    }

    return result;
}

export function parseRunProperty(elem: Element, props: RunProperties, xml: XmlParser) {
    if (parseCommonProperty(elem, props, xml))
        return true;

    return false;
}