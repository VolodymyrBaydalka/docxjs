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
    ignoreLastRenderedPageBreak: boolean;
}

export declare const defaultOptions: Options; 

export declare function renderAsync(data: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, options?: Partial<Options>): PromiseLike<any>;
