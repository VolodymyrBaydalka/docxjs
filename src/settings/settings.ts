import { DocumentParser } from "../document-parser";
import { Length } from "../document/common";
import { XmlParser } from "../parser/xml-parser";

export interface WmlSettings {
	defaultTabStop: Length;
	footnoteProps: NoteProperties;
	endnoteProps: NoteProperties;
	autoHyphenation: boolean;
}

export interface NoteProperties {
	nummeringFormat: string;
	defaultNoteIds: string[];
}

export function parseSettings(elem: Element, xml: XmlParser) {
	var result = {} as WmlSettings; 

	for (let el of xml.elements(elem)) {
		switch(el.localName) {
			case "defaultTabStop": result.defaultTabStop = xml.lengthAttr(el, "val"); break;
			case "footnotePr": result.footnoteProps = parseNoteProperties(el, xml); break;
			case "endnotePr": result.endnoteProps = parseNoteProperties(el, xml); break;
			case "autoHyphenation": result.autoHyphenation = xml.boolAttr(el, "val"); break;
		}
	}

    return result;
}

export function parseNoteProperties(elem: Element, xml: XmlParser) {
	var result = {
		defaultNoteIds: []
	} as NoteProperties; 

	for (let el of xml.elements(elem)) {
		switch(el.localName) {
			case "numFmt": 
				result.nummeringFormat = xml.attr(el, "val");
				break;

			case "footnote": 
			case "endnote": 
				result.defaultNoteIds.push(xml.attr(el, "id"));
				break;
		}
	}

    return result;
}