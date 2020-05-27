import { ElementBase } from "./element-base";
import { OpenXmlElement } from "../dom/dom";
import { updateTabStop } from "../javascript";
import { Paragraph } from "./paragraph";
import { RenderContext } from "../dom/render-context";
import { element } from "../parser/xml-serialize";

@element("tab")
export class Tab extends ElementBase {

    render(ctx: RenderContext): Node {
        var tabSpan = ctx.html.createElement("span");

        tabSpan.innerHTML = "&emsp;";//"&nbsp;";

        if (ctx.options.experimental) {
            setTimeout(() => {
                var paragraph = findParent<Paragraph>(this);

                if(paragraph.props.tabs == null)
                    return;

                paragraph.props.tabs.sort((a, b) => a.position.value - b.position.value);
                tabSpan.style.display = "inline-block";
                updateTabStop(tabSpan, paragraph.props.tabs);
            }, 0);
        }

        return tabSpan;
    }
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement): T {
    var parent = elem.parent;

    while (parent != null && !(parent instanceof Paragraph))
        parent = parent.parent;
    
    return <T>parent;
}