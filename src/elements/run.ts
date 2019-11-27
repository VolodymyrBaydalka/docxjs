import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export class Run extends ContainerBase {
    props: RunProeprties = {} as RunProeprties;

    //TODO
    fldCharType: any;
    instrText: any;
    href: any;

    render(ctx: RenderContext) : Node {
        if (this.fldCharType || this.instrText)
            return null;

        var elem = this.renderContainer(ctx, "span");
        var wrapper: HTMLElement = null;

        if(this.href)
        {
            wrapper = ctx.html.createElement("a");
            (wrapper as HTMLAnchorElement).href = this.href;
        }
        else
        {
            switch(this.props.verticalAlignment) {
                case "subscript": 
                    wrapper = ctx.html.createElement("sub");
                    break;

                case "superscript": 
                    wrapper = ctx.html.createElement("sup");
                    break;
            }
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