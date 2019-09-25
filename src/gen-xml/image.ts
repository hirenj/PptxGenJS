
import { genXmlPlaceholder } from './placeholder'

import { ISlideObject, ISlide, ISlideLayout } from '../core-interfaces'

import { encodeXmlEntities, getSmartParseNumber } from '../gen-utils'


let imageSizingXml = {
	cover: function(imgSize, boxDim) {
		var imgRatio = imgSize.h / imgSize.w,
			boxRatio = boxDim.h / boxDim.w,
			isBoxBased = boxRatio > imgRatio,
			width = isBoxBased ? boxDim.h / imgRatio : boxDim.w,
			height = isBoxBased ? boxDim.h : boxDim.w * imgRatio,
			hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
			vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
		return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '"/><a:stretch/>'
	},
	contain: function(imgSize, boxDim) {
		var imgRatio = imgSize.h / imgSize.w,
			boxRatio = boxDim.h / boxDim.w,
			widthBased = boxRatio > imgRatio,
			width = widthBased ? boxDim.w : boxDim.h / imgRatio,
			height = widthBased ? boxDim.w * imgRatio : boxDim.h,
			hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
			vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
		return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '"/><a:stretch/>'
	},
	crop: function(imageSize, boxDim) {
		var l = boxDim.x,
			r = imageSize.w - (boxDim.x + boxDim.w),
			t = boxDim.y,
			b = imageSize.h - (boxDim.y + boxDim.h),
			lPerc = Math.round(1e5 * (l / imageSize.w)),
			rPerc = Math.round(1e5 * (r / imageSize.w)),
			tPerc = Math.round(1e5 * (t / imageSize.h)),
			bPerc = Math.round(1e5 * (b / imageSize.h))
		return '<a:srcRect l="' + lPerc + '" r="' + rPerc + '" t="' + tPerc + '" b="' + bPerc + '"/><a:stretch/>'
	},
}

export function genXmlImage(slide: ISlide | ISlideLayout, slideItemObj: ISlideObject, idx: number, placeholderObj: ISlideObject,x,y,cx,cy,locationAttr: string): string {
	var sizing = slideItemObj.options.sizing,
		rounding = slideItemObj.options.rounding,
		width = cx,
		height = cy

	let strSlideXml = ''

	strSlideXml += '<p:pic>'
	strSlideXml += '  <p:nvPicPr>'
	strSlideXml += '    <p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '" descr="' + encodeXmlEntities(slideItemObj.image) + '">'
	if (slideItemObj.hyperlink && slideItemObj.hyperlink.url)
		strSlideXml +=
			'<a:hlinkClick r:id="rId' +
			slideItemObj.hyperlink.rId +
			'" tooltip="' +
			(slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') +
			'"/>'
	if (slideItemObj.hyperlink && slideItemObj.hyperlink.slide)
		strSlideXml +=
			'<a:hlinkClick r:id="rId' +
			slideItemObj.hyperlink.rId +
			'" tooltip="' +
			(slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') +
			'" action="ppaction://hlinksldjump"/>'
	strSlideXml += '    </p:cNvPr>'
	strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
	strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>'
	strSlideXml += '  </p:nvPicPr>'
	strSlideXml += '<p:blipFill>'
	// NOTE: This works for both cases: either `path` or `data` contains the SVG
	if (
		(slide['relsMedia'] || []).filter(rel => {
			return rel.rId == slideItemObj.imageRid
		})[0] &&
		(slide['relsMedia'] || []).filter(rel => {
			return rel.rId == slideItemObj.imageRid
		})[0]['extn'] == 'svg'
	) {
		strSlideXml += '<a:blip r:embed="rId' + (slideItemObj.imageRid - 1) + '">'
		strSlideXml += ' <a:extLst>'
		strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
		strSlideXml += '   <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId' + slideItemObj.imageRid + '"/>'
		strSlideXml += '  </a:ext>'
		strSlideXml += ' </a:extLst>'
		strSlideXml += '</a:blip>'
	} else {
		strSlideXml += '<a:blip r:embed="rId' + slideItemObj.imageRid + '"/>'
	}
	if (sizing && sizing.type) {
		var boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slide.presLayout) : cx,
			boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slide.presLayout) : cy,
			boxX = getSmartParseNumber(sizing.x || 0, 'X', slide.presLayout),
			boxY = getSmartParseNumber(sizing.y || 0, 'Y', slide.presLayout)

		strSlideXml += imageSizingXml[sizing.type]({ w: width, h: height }, { w: boxW, h: boxH, x: boxX, y: boxY })
		width = boxW
		height = boxH
	} else {
		strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>'
	}
	strSlideXml += '</p:blipFill>'
	strSlideXml += '<p:spPr>'
	strSlideXml += ' <a:xfrm' + locationAttr + '>'
	strSlideXml += '  <a:off x="' + x + '" y="' + y + '"/>'
	strSlideXml += '  <a:ext cx="' + width + '" cy="' + height + '"/>'
	strSlideXml += ' </a:xfrm>'
	strSlideXml += ' <a:prstGeom prst="' + (rounding ? 'ellipse' : 'rect') + '"><a:avLst/></a:prstGeom>'
	strSlideXml += '</p:spPr>'
	strSlideXml += '</p:pic>'
	return strSlideXml
}