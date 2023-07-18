export interface ChartElement {
	/**
	 * 标题
	 */
	title: string;
	/**
	 * 横坐标标题
	 */
	catAx: string;
	/**
	 * 纵坐标标题
	 */
	valAx: string;

	/**
	 * 图表数据
	 */
	chartList: Chart[];
}

/**
 * 图表数据
 */
export interface Chart {
	serList: Ser[];
}

export interface Ser {
	title: string;
	catList: string[];
	valList: string[];
}
