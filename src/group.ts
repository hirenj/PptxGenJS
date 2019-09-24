
import { SLIDE_OBJECT_TYPES, CHART_TYPE_NAMES } from './core-enums'
import {
	IChartMulti,
	IChartOpts,
	IImageOpts,
	ILayout,
	ISlide,
	IMediaOpts,
	ISlideLayout,
	ISlideNumber,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ISlideObject,
	IShape,
	IShapeOptions,
	ITableOptions,
	IText,
	ITextOpts,
	TableRow,
} from './core-interfaces'

import Slide from './slide'

import {
		addShapeDefinition,
		addChartDefinition,
		addGroupDefinition,
		addImageDefinition,
		addMediaDefinition,
		addTableDefinition,
		addTextDefinition	
	   } from './gen-objects'

export class Group {
	constructor(ownerSlide: Slide) {
		this.ownerSlide = ownerSlide
		this.data = [];
	}
	public ownerSlide: Slide
	public data: ISlideObject[]
	public type: SLIDE_OBJECT_TYPES = SLIDE_OBJECT_TYPES.group

	/**
	 * Generate the chart based on input data.
	 * @see OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
	 * @param {CHART_TYPE_NAMES|IChartMulti[]} `type` - chart type
	 * @param {object[]} data - a JSON object with follow the following format
	 * @param {IChartOpts} options - chart options
	 * @example
	 * {
	 *   title: 'eSurvey chart',
	 *   data: [
	 *		{
	 *			name: 'Income',
	 *			labels: ['2005', '2006', '2007', '2008', '2009'],
	 *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
	 *		},
	 *		{
	 *			name: 'Expense',
	 *			labels: ['2005', '2006', '2007', '2008', '2009'],
	 *			values: [18.1, 22.8, 23.9, 25.1, 25]
	 *		}
	 *	 ]
	 * }
	 * @return {Group} this class
	 */
	addChart(type: CHART_TYPE_NAMES | IChartMulti[], data: [], options?: IChartOpts): Group {
		addChartDefinition(this, type, data, options)
		return this
	}

	/**
	 * Add Group object
	 * @return {Group} Group class
	 */
	 addGroup(): Group {
		return addGroupDefinition(this);
	 }

	/**
	 * Add Image object
	 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
	 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
	 * @param {IImageOpts} options - image options
	 * @return {Group} this class
	 */
	addImage(options: IImageOpts): Group {
		addImageDefinition(this, options)
		return this
	}

	/**
	 * Add Media (audio/video) object
	 * @param {IMediaOpts} options - media options
	 * @return {Group} this class
	 */
	addMedia(options: IMediaOpts): Group {
		addMediaDefinition(this, options)
		return this
	}

	/**
	 * Add shape object to Slide
	 * @param {IShape} shape - shape object
	 * @param {IShapeOptions} options - shape options
	 * @return {Group} this class
	 */
	addShape(shape: IShape, options?: IShapeOptions): Group {
		addShapeDefinition(this, shape, options)
		return this
	}

	/**
	 * Add shape object to Slide
	 * @note can be recursive
	 * @param {TableRow[]} arrTabRows - table rows
	 * @param {ITableOptions} options - table options
	 * @return {Group} this class
	 */
	addTable(arrTabRows: TableRow[], options?: ITableOptions): Group {
		// FIXME: TODO-3: we pass `this` - we dont need to pass layouts - they can be read from this!
		addTableDefinition(this, arrTabRows, options, this.ownerSlide.slideLayout, this.ownerSlide.presLayout, this.ownerSlide.addSlide, this.ownerSlide.getSlide)
		return this
	}

	/**
	 * Add text object to Slide
	 * @param {string|IText[]} text - text string or complex object
	 * @param {ITextOpts} options - text options
	 * @return {Group} this class
	 * @since: 1.0.0
	 */
	addText(text: string | IText[], options?: ITextOpts): Group {
		addTextDefinition(this, text, options, false)
		return this
	}

}