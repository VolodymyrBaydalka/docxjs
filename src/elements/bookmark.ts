import { ElementBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export class Bookmark extends ElementBase {
    name: string;

    render(ctx: RenderContext) {
        var elem = ctx.html.createElement("span");
        elem.id = this.name;
        return elem;
    }
}