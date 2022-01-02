import { OpenXmlElement } from "./dom";

export interface WmlInstructionText extends OpenXmlElement {
    text: string;
}

export interface WmlFieldChar extends OpenXmlElement {
    charType: 'begin' | 'end' | 'separate' | string;
    lock: boolean;
}

export interface WmlFieldSimple extends OpenXmlElement {
    instruction: string;
    lock: boolean;
    dirty: boolean;
}