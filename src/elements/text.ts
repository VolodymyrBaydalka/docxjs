import { ElementBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export class Text extends ElementBase {
    text: string;

    render(context: RenderContext) {
        return context.html.createTextNode(this.text);
    }
}

export class Symbol extends ElementBase {
    font: string;
    char: string;

    render(ctx: RenderContext) {
        var span = ctx.html.createElement("span");
        span.style.fontFamily = this.font;
        span.innerHTML = `&#x${this.char};`
        return span;
    }
}