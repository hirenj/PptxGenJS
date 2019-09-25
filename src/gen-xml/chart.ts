
import { genXmlPlaceholder } from './placeholder'

import { ISlideObject, ISlide, ISlideLayout } from '../core-interfaces'

export function genXmlChart(slide: ISlide | ISlideLayout, slideItemObj: ISlideObject, idx: number, placeholderObj: ISlideObject,  x, y, cx, cy ): string {
	let strSlideXml = '<p:graphicFrame>'
	strSlideXml += ' <p:nvGraphicFramePr>'
	strSlideXml += '   <p:cNvPr id="' + (idx + 2) + '" name="Chart ' + (idx + 1) + '"/>'
	strSlideXml += '   <p:cNvGraphicFramePr/>'
	strSlideXml += '   <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>'
	strSlideXml += ' </p:nvGraphicFramePr>'
	strSlideXml += ' <p:xfrm>'
	strSlideXml += '  <a:off x="' + x + '" y="' + y + '"/>'
	strSlideXml += '  <a:ext cx="' + cx + '" cy="' + cy + '"/>'
	strSlideXml += ' </p:xfrm>'
	strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
	strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
	strSlideXml += '   <c:chart r:id="rId' + slideItemObj.chartRid + '" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
	strSlideXml += '  </a:graphicData>'
	strSlideXml += ' </a:graphic>'
	strSlideXml += '</p:graphicFrame>'
	return strSlideXml
}