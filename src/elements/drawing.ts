import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export class Drawing extends ContainerBase {
    render(ctx: RenderContext): Node {
        var elem = this.renderContainer(ctx, "div");

        elem.style.display = "inline-block";
        elem.style.position = "relative";
        elem.style.textIndent = "0px";

        return elem 
    }
}