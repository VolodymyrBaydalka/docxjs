import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { ChartElement } from './chart/chart';
import { IDomChart } from './document/dom';

export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
	renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
	useBase64URL: boolean;
	useMathMLPolyfill: boolean;
	renderChanges: boolean;
    /**
	 * 指定chart1渲染方法
	 * chart1: (chart: ChartElement) => IDomChart
	 * 
	 * 折线图渲染方法（非组合图表）
	 * lineChart: (chart: ChartElement) => IDomChart
	 * 
	 * 默认渲染方法（非组合图表可用）
	 * defaultRender: (chart: ChartElement) => IDomChart
	 * 
	 * 组合图表渲染方法
	 * mixedChart: (chart: ChartElement) => IDomChart
	 */
    renderCharts: Record<string, (chart: ChartElement) => IDomChart>;
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
	renderEndnotes: true,
	useBase64URL: false,
	useMathMLPolyfill: false,
	renderChanges: false,
    renderCharts: {},
}

export function praseAsync(data: Blob | any, userOptions: Partial<Options> = null): Promise<any>  {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer(window.document);
	const doc = await WordDocument.load(data, new DocumentParser(ops), ops)

	renderer.render(doc, bodyContainer, styleContainer, ops);
	
    return doc;
}