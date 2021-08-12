import { Part } from "../common/part";
import { ExtendedPropsDeclaration, parseExtendedProps } from "./props";

export class ExtendedPropsPart extends Part {
    props: ExtendedPropsDeclaration;

    parseXml(root: Element) {
        this.props = parseExtendedProps(root, this._package.xmlParser);
    }
}