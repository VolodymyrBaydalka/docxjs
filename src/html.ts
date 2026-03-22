import { isString } from "./utils";

export enum ns {
    html = "http://www.w3.org/1999/xhtml",
    svg = "http://www.w3.org/2000/svg",
    mathML = "http://www.w3.org/1998/Math/MathML"
}

export type HElement = {
    ns?: ns;
    tagName: "#fragment" | "#comment" | string;
    className?: string;
    style?: string | Record<string, string>;
    children?: (HElement | Node | string)[];
} & Record<string, any>;

export function h(elem: HElement | Node | string) {
    if (isString(elem)) return document.createTextNode(elem);
    if (elem instanceof Node) return elem;

    const { ns, tagName, className, style, children, ...props } = elem;

    if (tagName === "#fragment") return document.createDocumentFragment();
    if (tagName === "#comment") return document.createComment(children[0] as string);
    const result = (ns ? document.createElementNS(ns, tagName) : document.createElement(tagName)) as HTMLElement | SVGElement | MathMLElement;
    if (className) result.setAttribute("class", className);
    if (style) {
        if (isString(style)) {
            result.setAttribute("style", style);
        } else {
            Object.assign(result.style, style);
        }
    }
    if (props) {
        for (const [key, value] of Object.entries(props))
            if (value !== undefined)
                result[key] = value
    }
    if (children) children.forEach(c => result.appendChild(h(c)));
    return result;
}

export function cx(...classNames: (string | false | undefined)[]) {
    return classNames.filter(Boolean).join(" ");
}