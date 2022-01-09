import { Length } from "./document/common";
import { ParagraphTab } from "./document/paragraph";

const defaultTab: ParagraphTab = { position: { value: 0, type: "pt" }, leader: "none", style: "left" };
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
    const pOffset = pbb.left + marginLeft;
    const left = (ebb.left - pOffset) * pixelToPoint;
    const tab = tabs.find(t => t.style != "clear" && t.position.value > left);

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

		width = tab.position.value - offset * pixelToPoint;
    } else {
        width = tab.position.value - left;
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