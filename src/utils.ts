import { OpenXmlElement } from "./document/dom";

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

export function resolvePath(path: string, base: string): string {
    try {
        const prefix = "file://docx/";
        const url = new URL(path, prefix + base).toString();
        return url.substr(prefix.length);
    } catch {
        return `${base}${path}`;
    }
}

export function keyBy<T = any>(array: T[], by: (x: T) => any): Record<any, T> {
    return array.reduce((a, x) => {
        a[by(x)] = x;
        return a;
    }, {});
}

export function clone<T>(object: T): T {
    if(object === undefined) {
        return undefined;
    }
    if(object === null) {
        return null;
    }
    try {
        return JSON.parse(JSON.stringify(object));
    } catch(e) {
        console.warn(`Couldn't clone object:`, object);
        return object;
    }
}

export function isObject(item) {
    return (item && typeof item === 'object' && !Array.isArray(item));
}

export function mergeDeep(target, ...sources) {
    if (!sources.length) 
        return target;
    
    const source = sources.shift();

    if (isObject(target) && isObject(source)) {
        for (const key in source) {
            if (isObject(source[key])) {
                const val = target[key] ?? (target[key] = {});
                mergeDeep(val, source[key]);
            } else {
                target[key] = source[key];
            }
        }
    }

    return mergeDeep(target, ...sources);
}