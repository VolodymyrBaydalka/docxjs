import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import { keyBy } from "../utils";

export type CommentsExtended = {
    paraId: string;
    paraIdParent?: string;
    done: boolean;
}

export class CommentsExtendedPart extends Part {
    comments: CommentsExtended[] = [];
    commentMap: Record<string, CommentsExtended>;

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

	parseXml(root: Element) {
        const xml = this._package.xmlParser;

        for (let el of xml.elements(root, "commentEx")) {
            this.comments.push({
                paraId: xml.attr(el, 'paraId'),
                paraIdParent: xml.attr(el, 'paraIdParent'),
                done: xml.boolAttr(el, 'done')
            });
        }

		this.commentMap = keyBy(this.comments, x => x.paraId);
    }
}
