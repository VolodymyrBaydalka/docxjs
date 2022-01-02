import { XmlParser } from "../parser/xml-parser";

export interface CustomProperty {
	formatId: string;
	name: string;
	type: string;
	value: string;
}

export function parseCustomProps(root: Element, xml: XmlParser): CustomProperty[] {
	return xml.elements(root, "property").map(e => {
		const firstChild = e.firstChild;

		return {
			formatId: xml.attr(e, "fmtid"),
			name: xml.attr(e, "name"),
			type: firstChild.nodeName,
			value: firstChild.textContent
		};
	});
}