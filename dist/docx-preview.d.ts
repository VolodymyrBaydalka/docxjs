export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
    renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
    useBase64URL: boolean;
    renderChanges: boolean;
}

export declare const defaultOptions: Options;
export declare function praseAsync(data: Blob | any, userOptions: Partial<Options> | null): Promise<any>;
export declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement | null, options: Partial<Options> | null): Promise<any>;
