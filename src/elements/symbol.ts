import { ElementBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { fromAttribute } from "../parser/xml-serialize";

export class Symbol extends ElementBase {
    @fromAttribute("font")
    font: string;
    @fromAttribute("char")
    char: string;

    render(ctx: RenderContext) {
        var span = ctx.html.createElement("span");
        span.style.fontFamily = this.font;
        span.innerHTML = `&#x${this.char};`
        return span;
    }
}