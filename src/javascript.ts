import { ParagraphTab } from "./dom/paragraph";

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], pixelToPoint: number = 72 / 96) {

    let p = elem.closest("p");

    let tbb = elem.getBoundingClientRect();
    let pbb = p.getBoundingClientRect();

    let left = (tbb.left - pbb.left) * pixelToPoint;
    let tab = tabs.find(t => t.style != "clear" && t.position.value > left);

    if(tab == null)
        return;

    elem.style.display = "inline-block";
    elem.style.width = `${(tab.position.value - left)}pt`;    

    switch (tab.leader) {
        case "dot":
        case "middleDot":
            elem.style.borderBottom = "1px black dotted";
            break;

        case "hyphen":
        case "heavy":
        case "underscore":
            elem.style.borderBottom = "1px black solid";
            break;
    }
}