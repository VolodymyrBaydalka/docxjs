import { Part } from "../common/part";
import { CustomProperty, parseCustomProps } from "./custom-props";

export class CustomPropsPart extends Part {
    props: CustomProperty[];

    parseXml(root: Element) {
        this.props = parseCustomProps(root, this._package.xmlParser);
    }
}