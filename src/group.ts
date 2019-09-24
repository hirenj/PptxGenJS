
import { SLIDE_OBJECT_TYPES } from './core-enums'
import { ISlideObject, IShape, IShapeOptions } from './core-interfaces'

import { addShapeDefinition } from './gen-objects'

export class Group {
	constructor() {
		this.data = [];
	}
	public data: ISlideObject[]
	public type: SLIDE_OBJECT_TYPES = SLIDE_OBJECT_TYPES.group

	/**
	 * Add shape object to Group
	 * @param {IShape} shape - shape object
	 * @param {IShapeOptions} options - shape options
	 * @return {Group} this class
	 */
	addShape(shape: IShape, options?: IShapeOptions): Group {
		addShapeDefinition(this, shape, options)
		return this
	}

}