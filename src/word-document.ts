import { OutputType } from "jszip";

import { DocumentParser } from './document-parser';
import { Relationship, RelationshipTypes } from './common/relationship';
import { Part } from './common/part';
import { FontTablePart } from './font-table/font-table';
import { OpenXmlPackage } from './common/open-xml-package';
import { DocumentPart } from './document/document-part';
import { blobToBase64, resolvePath, splitPath } from './utils';
import { NumberingPart } from './numbering/numbering-part';
import { StylesPart } from './styles/styles-part';
import { FooterPart, HeaderPart } from "./header-footer/parts";
import { ExtendedPropsPart } from "./document-props/extended-props-part";
import { CorePropsPart } from "./document-props/core-props-part";
import { ThemePart } from "./theme/theme-part";
import { EndnotesPart, FootnotesPart } from "./notes/parts";
import { SettingsPart } from "./settings/settings-part";
import { CustomPropsPart } from "./document-props/custom-props-part";
import { CommentsPart } from "./comments/comments-part";
import { CommentsExtendedPart } from "./comments/comments-extended-part";

const topLevelRels = [
	{ type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
	{ type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
	{ type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
	{ type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
];

export class WordDocument {
	private _package: OpenXmlPackage;
	private _parser: DocumentParser;
	private _options: any;

	rels: Relationship[];
	parts: Part[] = [];
	partsMap: Record<string, Part> = {};

	documentPart: DocumentPart;
	fontTablePart: FontTablePart;
	numberingPart: NumberingPart;
	stylesPart: StylesPart;
	footnotesPart: FootnotesPart;
	endnotesPart: EndnotesPart;
	themePart: ThemePart;
	corePropsPart: CorePropsPart;
	extendedPropsPart: ExtendedPropsPart;
	settingsPart: SettingsPart;
	commentsPart: CommentsPart;
	commentsExtendedPart: CommentsExtendedPart;

	static async load(blob: Blob | any, parser: DocumentParser, options: any): Promise<WordDocument> {
		var d = new WordDocument();

		d._options = options;
		d._parser = parser;
		d._package = await OpenXmlPackage.load(blob, options);
		d.rels = await d._package.loadRelationships();

		await Promise.all(topLevelRels.map(rel => {
			const r = d.rels.find(x => x.type === rel.type) ?? rel; //fallback                    
			return d.loadRelationshipPart(r.target, r.type);
		}));

		return d;
	}

	save(type = "blob"): Promise<any> {
		return this._package.save(type);
	}

	private async loadRelationshipPart(path: string, type: string): Promise<Part> {
		if (this.partsMap[path])
			return this.partsMap[path];

		if (!this._package.get(path))
			return null;

		let part: Part = null;

		switch (type) {
			case RelationshipTypes.OfficeDocument:
				this.documentPart = part = new DocumentPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.FontTable:
				this.fontTablePart = part = new FontTablePart(this._package, path);
				break;

			case RelationshipTypes.Numbering:
				this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Styles:
				this.stylesPart = part = new StylesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Theme:
				this.themePart = part = new ThemePart(this._package, path);
				break;

			case RelationshipTypes.Footnotes:
				this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Endnotes:
				this.endnotesPart = part = new EndnotesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Footer:
				part = new FooterPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Header:
				part = new HeaderPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.CoreProperties:
				this.corePropsPart = part = new CorePropsPart(this._package, path);
				break;

			case RelationshipTypes.ExtendedProperties:
				this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
				break;

			case RelationshipTypes.CustomProperties:
				part = new CustomPropsPart(this._package, path);
				break;
	
			case RelationshipTypes.Settings:
				this.settingsPart = part = new SettingsPart(this._package, path);
				break;

			case RelationshipTypes.Comments:
				this.commentsPart = part = new CommentsPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.CommentsExtended:
				this.commentsExtendedPart = part = new CommentsExtendedPart(this._package, path);
				break;
		}

		if (part == null)
			return Promise.resolve(null);

		this.partsMap[path] = part;
		this.parts.push(part);

		await part.load();

		if (part.rels?.length > 0) {
			const [folder] = splitPath(part.path);
			await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
		}

		return part;
	}

	async loadDocumentImage(id: string, part?: Part): Promise<string> {
		const x = await this.loadResource(part ?? this.documentPart, id, "blob");
		return this.blobToURL(x);
	}

	async loadNumberingImage(id: string): Promise<string> {
		const x = await this.loadResource(this.numberingPart, id, "blob");
		return this.blobToURL(x);
	}

	async loadFont(id: string, key: string): Promise<string> {
		const x = await this.loadResource(this.fontTablePart, id, "uint8array");
		return x ? this.blobToURL(new Blob([deobfuscate(x, key)])) : x;
	}

	async loadAltChunk(id: string, part?: Part): Promise<string> {
		return await this.loadResource(part ?? this.documentPart, id, "string");
	}

	private blobToURL(blob: Blob): string | Promise<string> {
		if (!blob)
			return null;

		if (this._options.useBase64URL) {
			return blobToBase64(blob);
		}

		return URL.createObjectURL(blob);
	}

	findPartByRelId(id: string, basePart: Part = null) {
		var rel = (basePart.rels ?? this.rels).find(r => r.id == id);
		const folder = basePart ? splitPath(basePart.path)[0] : '';
		return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
	}

	getPathById(part: Part, id: string): string {
		const rel = part.rels.find(x => x.id == id);
		const [folder] = splitPath(part.path);
		return rel ? resolvePath(rel.target, folder) : null;
	}

	private loadResource(part: Part, id: string, outputType: OutputType) {
		const path = this.getPathById(part, id);
		return path ? this._package.load(path, outputType) : Promise.resolve(null);
	}
}

export function deobfuscate(data: Uint8Array, guidKey: string): Uint8Array {
	const len = 16;
	const trimmed = guidKey.replace(/{|}|-/g, "");
	const numbers = new Array(len);

	for (let i = 0; i < len; i++)
		numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);

	for (let i = 0; i < 32; i++)
		data[i] = data[i] ^ numbers[i % len]

	return data;
}