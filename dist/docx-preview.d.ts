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
//stub
export type WordDocument = any;
export declare const defaultOptions: Options;
export declare function praseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<WordDocument>;
export declare function renderDocument(document: WordDocument, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>);
export declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;
