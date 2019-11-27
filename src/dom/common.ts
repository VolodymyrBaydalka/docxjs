export const ns = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}

export type LengthType = "px" | "pt" | "%";

export interface Length {
    value: number;
    type: LengthType
}

export interface Border {
    color: string;
    type: string;
    size: Length;
}

export interface Borders {
    top: Border;
    left: Border;
    right: Border;
    botton: Border;
}

export interface Font {
    name: string;
    family: string;
}

export interface Column {
    space: Length;
    width: Length;
}

export interface Columns {
    space: Length;
    numberOfColumns: number;
    separator: boolean;
    equalWidth: boolean;
    columns: Column[];
}

export interface CommonProperties {
    fontSize: Length;
    color: string;
}

export function renderLength(l: Length): string {
    return !l ? null : `${l.value}${l.type}`;
}