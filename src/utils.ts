import { OpenXmlElement } from "./dom/dom";

export function addElementClass(element: OpenXmlElement, className: string): string {
    return element.className = appendClass(element.className, className);
}

export function appendClass(classList: string, className: string): string {
    return (!classList) ? className : `${classList} ${className}`
}

export function splitPath(path: string): [string, string] {
    let si = path.lastIndexOf('/') + 1;
    let folder = si == 0 ? "" : path.substring(0, si);
    let fileName = si == 0 ? path : path.substring(si);

    return [folder, fileName];
}

export function keyBy<T = any>(array: T[], by: (x: T) => any): Record<any, T> {
    return array.reduce((a, x) => {
        a[by(x)] = x;
        return a;
    }, {});
}