import { Part } from "../common/part";
import { CorePropsDeclaration, parseCoreProps } from "./core-props";

export class CorePropsPart extends Part {
    props: CorePropsDeclaration;

    parseXml(root: Element) {
        this.props = parseCoreProps(root, this._package.xmlParser);
    }
}