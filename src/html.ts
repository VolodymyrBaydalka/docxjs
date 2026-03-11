import { isString } from "./utils";

export enum ns {
	html = "http://www.w3.org/1999/xhtml",
	svg = "http://www.w3.org/2000/svg",
	mathML = "http://www.w3.org/1998/Math/MathML"
}

export type HElement = {
    ns?: ns;
    tagName: "#fragment" | "#comment" | string;
    classes?: string[];
    style?: Record<string, string>;
    children?: (HElement | Node | string)[];
} & Record<string, any>;

export function h(elem: HElement | Node | string): Node {
    if (isString(elem)) return document.createTextNode(elem);
    if (elem instanceof Node) return elem;

    const { ns, tagName, classes, style, children, ...props } = elem;

    if (tagName === "#fragment") return document.createDocumentFragment();
    if (tagName === "#comment") return document.createComment(children[0] as string);
    const result = (ns ? document.createElementNS(ns, tagName) : document.createElement(tagName)) as HTMLElement | SVGElement | MathMLElement;
    if (classes) result.classList.add(...classes.filter(Boolean));
    if (style) Object.assign(result.style, style);
    if (children) children.forEach(c => result.appendChild(h(c)));
    Object.assign(result, props);
    return result;
}
