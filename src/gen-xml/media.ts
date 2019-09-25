
import { ISlideObject, ISlide, ISlideLayout } from '../core-interfaces'

export function genXmlMedia(slide: ISlide | ISlideLayout, slideItemObj: ISlideObject,idx,x,y,cx,cy,locationAttr: string): string {
	let strSlideXml = ''
	if (slideItemObj.mtype == 'online') {
		strSlideXml += '<p:pic>'
		strSlideXml += ' <p:nvPicPr>'
		// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preview image rId, PowerPoint throws error!
		strSlideXml += ' <p:cNvPr id="' + (slideItemObj.mediaRid + 2) + '" name="Picture' + (idx + 1) + '"/>'
		strSlideXml += ' <p:cNvPicPr/>'
		strSlideXml += ' <p:nvPr>'
		strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>'
		strSlideXml += ' </p:nvPr>'
		strSlideXml += ' </p:nvPicPr>'
		// NOTE: `blip` is diferent than videos; also there's no preview "p:extLst" above but exists in videos
		strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' // NOTE: Preview image is required!
		strSlideXml += ' <p:spPr>'
		strSlideXml += '  <a:xfrm' + locationAttr + '>'
		strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>'
		strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>'
		strSlideXml += '  </a:xfrm>'
		strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
		strSlideXml += ' </p:spPr>'
		strSlideXml += '</p:pic>'
	} else {
		strSlideXml += '<p:pic>'
		strSlideXml += ' <p:nvPicPr>'
		// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
		strSlideXml +=
			' <p:cNvPr id="' +
			(slideItemObj.mediaRid + 2) +
			'" name="' +
			slideItemObj.media
				.split('/')
				.pop()
				.split('.')
				.shift() +
			'"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>'
		strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
		strSlideXml += ' <p:nvPr>'
		strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>'
		strSlideXml += '  <p:extLst>'
		strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
		strSlideXml += '    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/>'
		strSlideXml += '   </p:ext>'
		strSlideXml += '  </p:extLst>'
		strSlideXml += ' </p:nvPr>'
		strSlideXml += ' </p:nvPicPr>'
		strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 2) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' // NOTE: Preview image is required!
		strSlideXml += ' <p:spPr>'
		strSlideXml += '  <a:xfrm' + locationAttr + '>'
		strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>'
		strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>'
		strSlideXml += '  </a:xfrm>'
		strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
		strSlideXml += ' </p:spPr>'
		strSlideXml += '</p:pic>'
	}
	return strSlideXml
}