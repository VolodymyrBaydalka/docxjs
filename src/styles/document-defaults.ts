import { ParagraphProperties, parseParagraphProperties } from "../dom/paragraph";
import { RunProperties } from "../dom/run";
import { XmlParser } from "../parser/xml-parser";

export interface DocumentDefaults {
    paragraphProps: ParagraphProperties;
    runProps: RunProperties;

    paragraphStyle: any;
    runStyle: any;
}

export function parseDocumentDefaults(elem: Element, xml: XmlParser) {
    let result = <DocumentDefaults>{};

    for (let e of xml.elements(elem)) {
        switch(e.localName) {
            case "pPrDefault": 
                let pPrElem = xml.element(e, 'pPr');
                
                if (pPrElem)
                    result.paragraphProps = parseParagraphProperties(pPrElem, xml);
                break;

            case "rPrDefault":
                let rPrElem = xml.element(e, 'rPr');
                
                if (rPrElem)
                    result.runProps = parseParagraphProperties(rPrElem, xml);
                break;
        }
    }

    return result;
}