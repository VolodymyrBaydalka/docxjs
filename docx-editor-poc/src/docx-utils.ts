import { WmlParagraph } from "../../src/document/paragraph";
import { WmlRun } from "../../src/document/run";
import { WmlText } from "../../src/document/text";
import { getXmlElement, setXmlElement } from "./utils";

export function splitRuns(text: WmlText, index: number): WmlRun[] {
    const run = (text.parent as WmlRun);

    if (run.children.length === 1 && (index === 0 || index === text.text.length - 1))
        return [run];

    const paragraph = run.parent as WmlParagraph;
    const xRun = getXmlElement(run);

    const newRun = Object.assign(new WmlRun(), run); 
    const runIndex = paragraph.children.indexOf(run);
    const textIndex = newRun.children.indexOf(text);

    const xNewRun = xRun.cloneNode(true) as Element;
    xRun.after(xNewRun);

    setXmlElement(newRun, xNewRun);

    paragraph.children = paragraph.children.splice(runIndex, 0, newRun);

    return [run, newRun];
}