export interface ChartElement {
	/**
	 * Key
	 */
	key: string;

	/**
	 * Title
	 */
	title: string;

	/**
	 * X-axis title
	 */
	catAx: string;

	/**
	 * Y-axis title
	 */
	valAx: string;

	/**
	 * Chart data
	 */
	chartList: Chart[];
}

/**
 * Chart data
 */
export interface Chart {
	/**
	 * Chart type
	 */
	type: string;

	serList: Ser[];
}

export interface Ser {
	/**
	 * Series title
	 */
	title: string;

	/**
	 * Category list
	 */
	catList: string[];

	/**
	 * Value list
	 */
	valList: string[];
}
