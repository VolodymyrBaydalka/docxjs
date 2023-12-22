import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import { DocumentParser } from "../document-parser";
import { WmlComment } from "./elements";
import { keyBy } from "../utils";

export class CommentsPart extends Part {
    protected _documentParser: DocumentParser;

    comments: WmlComment[]
    commentMap: Record<string, WmlComment>;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

	parseXml(root: Element) {
        this.comments = this._documentParser.parseComments(root);
		this.commentMap = keyBy(this.comments, x => x.id);
    }
}