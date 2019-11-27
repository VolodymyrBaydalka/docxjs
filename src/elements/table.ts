import { ContainerBase } from "./element-base";
import { Length, renderLength } from "../dom/common";
import { RenderContext } from "../dom/render-context";

export class Table extends ContainerBase {

    columns: TableColumn[];

    render(ctx: RenderContext): Node {
        var elem = this.renderContainer(ctx, "table");

        if(this.columns)
            elem.appendChild(this.renderTableColumns(ctx, this.columns));

        return elem;
    }

    private renderTableColumns(ctx: RenderContext, columns: TableColumn[]) {
        let result = ctx.html.createElement("colGroup");

        for (let col of columns) {
            let colElem = ctx.html.createElement("col");

            if (col.width)
                colElem.style.width = renderLength(col.width);

            result.appendChild(colElem);
        }

        return result;
    }
}

export interface TableColumn {
    width: Length;
}