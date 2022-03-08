import { Length } from "./document/common";
import { ParagraphTab } from "./document/paragraph";

interface TabStop {
	pos: number;
	leader: string;
	style: string;
}

const defaultTab: TabStop = { pos: 0, leader: "none", style: "left" };
const maxTabs = 50;

export function computePixelToPoint(container: HTMLElement = document.body) {
	const temp = document.createElement("div");
	temp.style.width = '100pt';
	
	container.appendChild(temp);
	const result = 100 / temp.offsetWidth;
	container.removeChild(temp);

	return result
}

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length, pixelToPoint: number = 72 / 96) {
    const p = elem.closest("p");

    const ebb = elem.getBoundingClientRect();
    const pbb = p.getBoundingClientRect();
    const pcs = getComputedStyle(p);

	const tabStops = tabs?.length > 0 ? tabs.map(t => ({
		pos: lengthToPoint(t.position),
		leader: t.leader,
		style: t.style
	})).sort((a, b) => a.pos - b.pos) : [defaultTab];

	const lastTab = tabStops[tabStops.length - 1];
	const pWidthPt = pbb.width * pixelToPoint;
	const size = lengthToPoint(defaultTabSize);
    let pos = lastTab.pos + size;

    if (pos < pWidthPt) {
        for (; pos < pWidthPt && tabStops.length < maxTabs; pos += size) {
            tabStops.push({ ...defaultTab, pos: pos });
        }
    }

    const marginLeft = parseFloat(pcs.marginLeft);
    const pOffset = pbb.left + marginLeft;
    const left = (ebb.left - pOffset) * pixelToPoint;
    const tab = tabStops.find(t => t.style != "clear" && t.pos > left);

    if(tab == null)
        return;

    let width: number = 1;

    if (tab.style == "right" || tab.style == "center") {
		const tabStops = Array.from(p.querySelectorAll(`.${elem.className}`));
		const nextIdx = tabStops.indexOf(elem) + 1;
        const range = document.createRange();
        range.setStart(elem, 1);

		if (nextIdx < tabStops.length) {
			range.setEndBefore(tabStops[nextIdx]);
		} else {
			range.setEndAfter(p);
		}

		const mul = tab.style == "center" ? 0.5 : 1;
        const nextBB = range.getBoundingClientRect();
		const offset = nextBB.left + mul * nextBB.width - (pbb.left - marginLeft);

		width = tab.pos - offset * pixelToPoint;
    } else {
        width = tab.pos - left;
    }

    elem.innerHTML = "&nbsp;";
    elem.style.textDecoration = "inherit";
    elem.style.wordSpacing = `${width.toFixed(0)}pt`;

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

function lengthToPoint(length: Length): number {
	return parseFloat(length);
}