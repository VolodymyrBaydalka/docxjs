import { ElementBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { fromAttribute } from "../parser/xml-serialize";

export class BookmarkStart extends ElementBase {
    @fromAttribute("id")
    id: string;
    @fromAttribute("name")
    name: string;

    render(ctx: RenderContext) {
        var elem = ctx.html.createElement("span");
        elem.id = this.name;
        return elem;
    }
}

export class BookmarkEnd extends ElementBase {
    @fromAttribute("id")
    id: string;
}