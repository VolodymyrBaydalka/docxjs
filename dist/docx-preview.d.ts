export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    className: string;
}

export declare function renderAsync(data: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, options?: Partial<Options>): PromiseLike<any>;
