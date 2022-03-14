import { ParagraphProperties } from "./paragraph";
import { RunProperties } from "./run";

export interface IDomStyle {
    id: string;
    name?: string;
    cssName?: string;
    aliases?: string[];
    target: string;
    basedOn?: string;
    isDefault?: boolean;
    styles: IDomSubStyle[];
    linked?: string;
    next?: string;

    paragraphProps: ParagraphProperties;
    runProps: RunProperties;
}

export interface IDomSubStyle {
    target: string;
	mod?: string;
    values: Record<string, string>;
}