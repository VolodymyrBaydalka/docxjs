import { Part } from "../common/part";
import { ExtendedPropsDeclaration, parseExtendedProps } from "./extended-props";

export class ExtendedPropsPart extends Part {
    props: ExtendedPropsDeclaration;

    parseXml(root: Element) {
        this.props = parseExtendedProps(root, this._package.xmlParser);
    }
}