import { ElementBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { fromAttribute } from "../parser/xml-serialize";

export class Bookmark extends ElementBase {
    @fromAttribute("name")
    name: string;

    render(ctx: RenderContext) {
        var elem = ctx.html.createElement("span");
        elem.id = this.name;
        return elem;
    }
}