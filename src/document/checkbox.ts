import { DomType, OpenXmlElement } from "./dom";
import xml from '../parser/xml-parser';

export interface WmlCheckboxFormField extends OpenXmlElement {
	name: string,
	checked: boolean,
}

const checkboxDefaultName = "unknownCheckbox";

const getName = (checkboxElement: Element) => {
	const parentElement = checkboxElement.parentElement;
	if (!parentElement) return checkboxDefaultName;
	const statusTextElement = parentElement.getElementsByTagName("w:statusText")[0];
	const statusTextName = statusTextElement ? xml.attr(statusTextElement, "val") : "";
	if (statusTextName) return statusTextName;

	const nameElement = parentElement.getElementsByTagName("w:name")[0];
	const nameNodeText = nameElement ? xml.attr(nameElement, "val") : "";
	return nameNodeText || checkboxDefaultName;
}

const getDefaultChecked = (checkboxElement: Element) => {
	const defaultElement = checkboxElement.getElementsByTagName("w:default")[0];
	return defaultElement ? xml.boolAttr(defaultElement, "val", false) : false;
}

const parseCheckbox = (element: Element): WmlCheckboxFormField | null => {
	const checkboxElement = element.getElementsByTagName("w:checkBox")[0];
	if (!checkboxElement) return null;

	return {
		type: DomType.CheckboxFormField,
		name: getName(checkboxElement),
		checked: getDefaultChecked(checkboxElement),
	}
}

export { parseCheckbox }
