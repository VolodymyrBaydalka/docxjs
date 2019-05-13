import { OpenXmlElement } from "./dom/dom";

export function addElementClass(element: OpenXmlElement, className: string): string {
    return element.className = appendClass(element.className, className);
}

export function appendClass(classList: string, className: string): string {
    return (!classList) ? className : `${classList} ${className}`
}