import { ElementBase, renderStyleValues } from "./element-base"
import { RenderContext } from "../dom/render-context";

export class Image extends ElementBase {

    src: string;
    style: any = {};

    render(ctx: RenderContext): Node {
        let result = ctx.html.createElement("img");

        renderStyleValues(this.style, result);

        if (ctx.document) {
            ctx.document.loadDocumentImage(this.src).then(x => {
                result.src = x;
            });
        }

        return result;
    }
}