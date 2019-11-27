import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export class Run extends ContainerBase {
    render(ctx: RenderContext) : Node {
        return this.renderContainer(ctx, "span");
    }
}