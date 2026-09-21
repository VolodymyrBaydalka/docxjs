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

/**
 * Applying a tab widens its container, and every measurement taken afterwards is
 * taken against that changed layout: inside a table cell the column grows, more
 * default stops fit into it, and the next tab grows wider still. So all the tabs
 * are measured against the untouched layout first and only then applied.
 */
export function updateTabStops(tabs: { span: HTMLElement, stops: ParagraphTab[] }[], defaultTabSize: Length, pixelToPoint: number = 72 / 96) {
	const measured = tabs.map(t => measureTabStop(t.span, t.stops, defaultTabSize, pixelToPoint));

	tabs.forEach((t, i) => measured[i] != null && applyTabStop(t.span, measured[i]));
}

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length, pixelToPoint: number = 72 / 96) {
	const measured = measureTabStop(elem, tabs, defaultTabSize, pixelToPoint);

	if (measured != null)
		applyTabStop(elem, measured);
}

function applyTabStop(elem: HTMLElement, { width, leader }: { width: number, leader: string }) {
    elem.innerHTML = "&nbsp;";
    elem.style.textDecoration = "inherit";
    elem.style.wordSpacing = `${width.toFixed(0)}pt`;

    switch (leader) {
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

function measureTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length, pixelToPoint: number) {
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
        return null;

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

    // a tab never reaches past the right edge of its own paragraph; without this
    // a stop meant for the page width stretches the table cell that inherited it
    return { width: Math.min(width, pWidthPt - left), leader: tab.leader };
}

function lengthToPoint(length: Length): number {
	return parseFloat(length);
}