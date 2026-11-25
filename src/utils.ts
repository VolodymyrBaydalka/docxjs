export function escapeClassName(className: string) {
	return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
}

export function encloseFontFamily(fontFamily: string): string {
    return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
}

const serifFonts = /cambria|times|georgia|garamond|palatino|book antiqua|bookman|constantia|baskerville|century/i;
const monospaceFonts = /courier|consolas|menlo|monaco|lucida console|lucida sans typewriter/i;
const symbolFonts = /symbol|wingdings|webdings/i;

// Generic family to fall back on when a document font is not installed;
// without one the browser uses its default font (usually Times) for any font.
export function genericFontFamily(fontFamily: string): string {
    if (!fontFamily || symbolFonts.test(fontFamily)) return null;
    if (monospaceFonts.test(fontFamily)) return "monospace";
    if (serifFonts.test(fontFamily)) return "serif";
    return "sans-serif";
}

/** The font name followed by its generic fallback, for use in a font-family list. */
export function withGenericFontFamily(fontFamily: string): string {
    const generic = genericFontFamily(fontFamily);
    return generic ? `${encloseFontFamily(fontFamily)}, ${generic}` : fontFamily;
}

export function splitPath(path: string): [string, string] {
    let si = path.lastIndexOf('/') + 1;
    let folder = si == 0 ? "" : path.substring(0, si);
    let fileName = si == 0 ? path : path.substring(si);

    return [folder, fileName];
}

export function resolvePath(path: string, base: string): string {
    try {
        const prefix = "http://docx/";
        const url = new URL(path, prefix + base).toString();
        return url.substring(prefix.length);
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

export function blobToBase64(blob: Blob): Promise<string> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onloadend = () => resolve(reader.result as string);
		reader.onerror = () => reject();
		reader.readAsDataURL(blob);
	});
}

export function isObject(item) {
    return item && typeof item === 'object' && !Array.isArray(item);
}

export function isString(item: unknown): item is string {
    return typeof item === 'string' || item instanceof String;
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

export function parseCssRules(text: string): Record<string, string> {
	const result: Record<string, string> = {};

	for (const rule of text.split(';')) {
		const [key, val] = rule.split(':');
		result[key] = val;
	}

	return result
}

export function formatCssRules(style: Record<string, string>): string {
	return Object.entries(style).map((k, v) => `${k}: ${v}`).join(';');
}

export function asArray<T>(val: T | T[]): T[] {
	return Array.isArray(val) ? val : [val];
}

export function clamp(val, min, max) {
    return min > val ? min : (max < val ? max : val);
}