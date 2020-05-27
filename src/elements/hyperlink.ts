import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { fromAttribute, element, children } from "../parser/xml-serialize";
import { Run } from "./run";

@element("hyperlink")
//@children(Run)
export class Hyperlink extends ContainerBase {
    @fromAttribute("anchor")
    anchor: string;

    render(ctx: RenderContext): Node {
        var a = this.renderContainer(ctx, "a");

        if(this.anchor)
            a.href = `#${this.anchor}`;
        
        return a;
    }
}