import { ParagraphTab } from "./document/paragraph";

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], pixelToPoint: number = 72 / 96) {

    const p = elem.closest("p");

    const tbb = elem.getBoundingClientRect();
    const pbb = p.getBoundingClientRect();
    const pcs = getComputedStyle(p);

    const marginLeft = parseFloat(pcs.marginLeft);
    const textIntent = parseFloat(pcs.textIndent);
    const pOffset = pbb.left + marginLeft;
    let left = (tbb.left - pOffset) * pixelToPoint;
    let tab = tabs.find(t => t.style != "clear" && t.position.value > left);

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

export function updateDefaultTabStop(elem: HTMLElement, tabDefaultPtWidth: number, iterations: number = 3) {
    const pixelToPoint = 72 / 96;
    const p = elem.closest("p");
    const art = elem.closest("article");
    const pMarginLeft = parseFloat(p.style.marginLeft);
    const abb = art.getBoundingClientRect();
    const pbb = p.getBoundingClientRect();
    const tbb = elem.getBoundingClientRect();
    let tabLeft = (tbb.x - pbb.x) * pixelToPoint;

    let nextTabStopPosition: number = tabDefaultPtWidth;
    if(!Number.isNaN(pMarginLeft) && tabLeft < 0) {
        tabLeft = (tbb.x - abb.x) * pixelToPoint;
        nextTabStopPosition = (pbb.x - abb.x) * pixelToPoint;
    } else {
        if(tabLeft < 0 || tabDefaultPtWidth < 0) {
            return;
        }
        // +1 to avoid rounding errors.
        while (nextTabStopPosition < tabLeft + 1) {
            nextTabStopPosition += tabDefaultPtWidth;
        }
    }
    const desiredWidth: number = nextTabStopPosition - tabLeft;
    elem.style.display = "inline-block";
    elem.style.width = `${desiredWidth}pt`;

    // Changing the positions we might have to recalculate these one more time
    if(iterations-- > 0) {
        setTimeout(() => {
            updateDefaultTabStop(elem, tabDefaultPtWidth, iterations);
        }, 25);
    }
}