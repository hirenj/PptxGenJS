
import {
	BULLET_TYPES,
	CRLF,
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_PT,
	EMU,
	LAYOUT_IDX_SERIES_BASE,
	ONEPT,
	PLACEHOLDER_TYPES,
	SLDNUMFLDID,
	SLIDE_OBJECT_TYPES,
	DEF_PRES_LAYOUT_NAME,
} from '../core-enums'

import {
	ILayout,
	IShadowOptions,
	ISlide,
	IGroup,
	ISlideLayout,
	ISlideObject,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ITableCell,
	ITableCellOpts,
	IObjectOptions,
	IText,
	ITextOpts,
} from '../core-interfaces'

import { encodeXmlEntities, inch2Emu, genXmlColorSelection, getSmartParseNumber, convertRotationDegrees } from '../gen-utils'

import { genXmlChart } from './chart'
import { genXmlImage } from './image'
import { genXmlMedia } from './media'
import { genXmlTable } from './table'
import { genXmlText } from './text'

export function genXmlSlideObject(slide: ISlide | ISlideLayout, slideItemObj: ISlideObject, idx: number): String {
	let x = 0,
		y = 0,
		cx = getSmartParseNumber('75%', 'X', slide.presLayout),
		cy = 0
	let placeholderObj: ISlideObject
	let locationAttr = ''
	let strSlideXml = ''


	if ((slide as ISlide).slideLayout !== undefined && (slide as ISlide).slideLayout.rootGroup.data !== undefined && slideItemObj.options && slideItemObj.options.placeholder) {
		placeholderObj = slide['slideLayout']['rootGroup']['data'].filter((object: ISlideObject) => {
			return object.options.placeholder == slideItemObj.options.placeholder
		})[0]
	}

	// A: Set option vars
	slideItemObj.options = slideItemObj.options || {}

	if (typeof slideItemObj.options.x !== 'undefined') x = getSmartParseNumber(slideItemObj.options.x, 'X', slide.presLayout)
	if (typeof slideItemObj.options.y !== 'undefined') y = getSmartParseNumber(slideItemObj.options.y, 'Y', slide.presLayout)
	if (typeof slideItemObj.options.w !== 'undefined') cx = getSmartParseNumber(slideItemObj.options.w, 'X', slide.presLayout)
	if (typeof slideItemObj.options.h !== 'undefined') cy = getSmartParseNumber(slideItemObj.options.h, 'Y', slide.presLayout)

	// If using a placeholder then inherit it's position
	if (placeholderObj) {
		if (placeholderObj.options.x || placeholderObj.options.x == 0) x = getSmartParseNumber(placeholderObj.options.x, 'X', slide.presLayout)
		if (placeholderObj.options.y || placeholderObj.options.y == 0) y = getSmartParseNumber(placeholderObj.options.y, 'Y', slide.presLayout)
		if (placeholderObj.options.w || placeholderObj.options.w == 0) cx = getSmartParseNumber(placeholderObj.options.w, 'X', slide.presLayout)
		if (placeholderObj.options.h || placeholderObj.options.h == 0) cy = getSmartParseNumber(placeholderObj.options.h, 'Y', slide.presLayout)
	}
	//
	if (slideItemObj.options.flipH) locationAttr += ' flipH="1"'
	if (slideItemObj.options.flipV) locationAttr += ' flipV="1"'
	if (slideItemObj.options.rotate) locationAttr += ' rot="' + convertRotationDegrees(slideItemObj.options.rotate) + '"'

	// B: Add OBJECT to the current Slide
	switch (slideItemObj.type) {
		case SLIDE_OBJECT_TYPES.group:
			strSlideXml += '<p:grpSp>'
			strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
			strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
			strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
			slideItemObj.data.map( genXmlSlideObject.bind(null,slide) ).forEach( xml => strSlideXml += xml )
			strSlideXml += '</p:grpSp>'
			break;
		case SLIDE_OBJECT_TYPES.table:
			strSlideXml += genXmlTable(slide,slideItemObj,x,y,cx,cy)
			break
		case SLIDE_OBJECT_TYPES.text:
		case SLIDE_OBJECT_TYPES.placeholder:
			strSlideXml += genXmlText(slide,slideItemObj,idx,placeholderObj,x,y,cx,cy,locationAttr)
			break
		case SLIDE_OBJECT_TYPES.image:
			strSlideXml += genXmlImage(slide,slideItemObj,idx,placeholderObj,x,y,cx,cy,locationAttr)
			break
		case SLIDE_OBJECT_TYPES.media:
			strSlideXml += genXmlMedia(slide,slideItemObj,idx,x,y,cx,cy,locationAttr)
			break
		case SLIDE_OBJECT_TYPES.chart:
			strSlideXml += genXmlChart(slide,slideItemObj,idx,placeholderObj,x,y,cx,cy)
			break
	}

	return strSlideXml
}