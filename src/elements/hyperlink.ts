import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { fromAttribute } from "../parser/xml-serialize";

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