import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { element, children } from "../parser/xml-serialize";
import { Break } from "./break";
import { Symbol, Text } from "./text";
import { Tab } from "./tab";

@element("r")
@children(Text, Break, Symbol, Tab)
export class Run extends ContainerBase {
    props: RunProeprties = {} as RunProeprties;

    //TODO
    fldCharType: any;

    render(ctx: RenderContext) : Node {
        if (this.fldCharType)
            return null;

        var elem = this.renderContainer(ctx, "span");
        var wrapper: HTMLElement = null;

        switch(this.props.verticalAlignment) {
            case "subscript": 
                wrapper = ctx.html.createElement("sub");
                break;

            case "superscript": 
                wrapper = ctx.html.createElement("sup");
                break;
        }

        if(wrapper == null)
            return elem;
            
        wrapper.appendChild(elem);

        return wrapper;
    }
}

export type RunVerticalAligmentType = "subscript" | "superscript";

export interface RunProeprties {
    verticalAlignment: RunVerticalAligmentType | string; 
}