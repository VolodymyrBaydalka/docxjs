import { ElementBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export type BreakType =  "page" | "lastRenderedPageBreak" | "textWrapping";

export class Break extends ElementBase {
    break: BreakType | string;

    render(ctx: RenderContext) {
        return this.break == "textWrapping" ? ctx.html.createElement("br") : null;
    }
}
