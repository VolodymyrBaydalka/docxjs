export const namespaces = {
    wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}

export interface Length {
    value: number;
    type: "px" | "pt" | "%"
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