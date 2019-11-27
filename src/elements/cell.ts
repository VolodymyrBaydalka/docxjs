import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export class Cell extends ContainerBase {
    props: CellProperties = {} as CellProperties;
    
    render(ctx: RenderContext): Node {
        var elem = this.renderContainer(ctx, "td");

        if(this.props.gridSpan)
            elem.colSpan = this.props.gridSpan;

        return elem;
    }
}

export interface CellProperties {
    gridSpan: number;
}