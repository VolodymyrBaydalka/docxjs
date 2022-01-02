import { Length } from "./document/common";
import { ParagraphTab } from "./document/paragraph";

const defaultTab: ParagraphTab = { position: { value: 0, type: "pt" }, leader: "none", style: "left" };
const maxTabs = 50;

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length, pixelToPoint: number = 72 / 96) {

    const p = elem.closest("p");

    const tbb = elem.getBoundingClientRect();
    const pbb = p.getBoundingClientRect();
    const pcs = getComputedStyle(p);

	tabs = tabs && tabs.length > 0 ? tabs.sort((a, b) => a.position.value - b.position.value) : [defaultTab];

	const lastTab = tabs[tabs.length - 1];
	const pWidthPt = pbb.width * pixelToPoint;
	const size = defaultTabSize.value;
    let pos = lastTab.position.value + defaultTabSize.value;

    if (pos < pWidthPt) {
        tabs = [...tabs];

        for (; pos < pWidthPt && tabs.length < maxTabs; pos += size) {
            tabs.push({ ...defaultTab, position: { value: pos, type: "pt" } });
        }
    }

    const marginLeft = parseFloat(pcs.marginLeft);
    const textIntent = parseFloat(pcs.textIndent);
    const pOffset = pbb.left + marginLeft;
    const left = (tbb.left - pOffset) * pixelToPoint;
    const tab = tabs.find(t => t.style != "clear" && t.position.value > left);

    if(tab == null)
        return;

    let width: any = 1;

    if (tab.style == "right") {
        const range = document.createRange();
        range.setStart(p.firstChild, 0);
        range.setEndAfter(p);

        const nextBB = range.getBoundingClientRect();
        const prevRight = (nextBB.width + marginLeft + textIntent) * pixelToPoint;
        width = `${Math.floor(tab.position.value - prevRight)}pt`;
    } else {
        width = `${(tab.position.value - left)}pt`;
    }

    elem.innerHTML = "&nbsp;";
    elem.style.textDecoration = "inherit";
    elem.style.wordSpacing = width;

    switch (tab.leader) {
        case "dot":
        case "middleDot":
            elem.style.textDecoration = "underline";
            elem.style.textDecorationStyle = "dotted";
            break;

        case "hyphen":
        case "heavy":
        case "underscore":
            elem.style.textDecoration = "underline";
            break;
    }
}