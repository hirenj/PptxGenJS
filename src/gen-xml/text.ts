
import { genXmlPlaceholder } from './placeholder'

import { PowerPointShapes } from '../core-shapes'

import { ISlideObject, ISlide, ISlideLayout, ITableCell, IObjectOptions, IText, ITextOpts } from '../core-interfaces'

import {
	BULLET_TYPES,
	CRLF,
	EMU,
	ONEPT,
	SLIDE_OBJECT_TYPES,
} from '../core-enums'

import { encodeXmlEntities, genXmlColorSelection } from '../gen-utils'


function getShapeInfo(shapeName) {
	if (!shapeName) return PowerPointShapes.RECTANGLE

	if (typeof shapeName == 'object' && shapeName.name && shapeName.displayName && shapeName.avLst) return shapeName

	if (PowerPointShapes[shapeName]) return PowerPointShapes[shapeName]

	var objShape = Object.keys(PowerPointShapes).filter((key: string) => {
		return PowerPointShapes[key].name == shapeName || PowerPointShapes[key].displayName
	})[0]
	if (typeof objShape !== 'undefined' && objShape != null) return objShape

	return PowerPointShapes.RECTANGLE
}

/**
 * Builds `<a:bodyPr></a:bodyPr>` tag for "genXmlTextBody()"
 * @param {ISlideObject | ITableCell} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties(slideObject: ISlideObject | ITableCell): string {
	let bodyProperties = '<a:bodyPr'

	if (slideObject && slideObject.type === SLIDE_OBJECT_TYPES.text && slideObject.options.bodyProp) {
		// PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>

		// A: Enable or disable textwrapping none or square
		bodyProperties += slideObject.options.bodyProp.wrap ? ' wrap="' + slideObject.options.bodyProp.wrap + '"' : ' wrap="square"'

		// B: Textbox margins [padding]
		if (slideObject.options.bodyProp.lIns || slideObject.options.bodyProp.lIns == 0) bodyProperties += ' lIns="' + slideObject.options.bodyProp.lIns + '"'
		if (slideObject.options.bodyProp.tIns || slideObject.options.bodyProp.tIns == 0) bodyProperties += ' tIns="' + slideObject.options.bodyProp.tIns + '"'
		if (slideObject.options.bodyProp.rIns || slideObject.options.bodyProp.rIns == 0) bodyProperties += ' rIns="' + slideObject.options.bodyProp.rIns + '"'
		if (slideObject.options.bodyProp.bIns || slideObject.options.bodyProp.bIns == 0) bodyProperties += ' bIns="' + slideObject.options.bodyProp.bIns + '"'

		// C: Add rtl after margins
		bodyProperties += ' rtlCol="0"'

		// D: Add anchorPoints
		if (slideObject.options.bodyProp.anchor) bodyProperties += ' anchor="' + slideObject.options.bodyProp.anchor + '"' // VALS: [t,ctr,b]
		if (slideObject.options.bodyProp.vert) bodyProperties += ' vert="' + slideObject.options.bodyProp.vert + '"' // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		// E: Close <a:bodyPr element
		bodyProperties += '>'

		// F: NEW: Add autofit type tags
		if (slideObject.options.shrinkText) bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>' // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
		// MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
		// NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
		bodyProperties += slideObject.options.bodyProp.autoFit !== false ? '<a:spAutoFit/>' : ''

		// LAST: Close bodyProp
		bodyProperties += '</a:bodyPr>'
	} else {
		// DEFAULT:
		bodyProperties += ' wrap="square" rtlCol="0">'
		bodyProperties += '</a:bodyPr>'
	}

	// LAST: Return Close bodyProp
	return slideObject.type == SLIDE_OBJECT_TYPES.tablecell ? '<a:bodyPr/>' : bodyProperties
}

/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragprah-level formatting device
 * @param {ISlideObject|ITableCell} slideObj - slideObj -OR- table `cell` object
 * @returns XML containing the param object's text and formatting
 */
export function genXmlTextBody(slideObj: ISlideObject | ITableCell): string {
	let opts: IObjectOptions = slideObj.options || {}
	// FIRST: Shapes without text, etc. may be sent here during build, but have no text to render so return an empty string
	if (opts && slideObj.type != SLIDE_OBJECT_TYPES.tablecell && (typeof slideObj.text === 'undefined' || slideObj.text == null)) return ''

	// Vars
	let arrTextObjects: IText[] = []
	let tagStart = slideObj.type == SLIDE_OBJECT_TYPES.tablecell ? '<a:txBody>' : '<p:txBody>'
	let tagClose = slideObj.type == SLIDE_OBJECT_TYPES.tablecell ? '</a:txBody>' : '</p:txBody>'
	let strSlideXml = tagStart

	// STEP 1: Modify slideObj to be consistent array of `{ text:'', options:{} }`
	/* CASES:
		addText( 'string' )
		addText( 'line1\n line2' )
		addText( ['barry','allen'] )
		addText( [{text'word1'}, {text:'word2'}] )
		addText( [{text'line1\n line2'}, {text:'end word'}] )
	*/
	// A: Transform string/number into complex object
	if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
		slideObj.text = [{ text: slideObj.text.toString(), options: opts || {} }]
	}

	// STEP 2: Grab options, format line-breaks, etc.
	if (Array.isArray(slideObj.text)) {
		slideObj.text.forEach((obj, idx) => {
			// A: Set options
			obj.options = obj.options || opts || {}
			if (idx == 0 && obj.options && !obj.options.bullet && opts.bullet) obj.options.bullet = opts.bullet

			// B: Cast to text-object and fix line-breaks (if needed)
			if (typeof obj.text === 'string' || typeof obj.text === 'number') {
				// 1: Convert "\n" or any variation into CRLF
				obj.text = obj.text.toString().replace(/\r*\n/g, CRLF)

				// 2: Handle strings that contain "\n"
				if (obj.text.indexOf(CRLF) > -1) {
					// Remove trailing linebreak (if any) so the "if" below doesnt create a double CRLF+CRLF line ending!
					obj.text = obj.text.replace(/\r\n$/g, '')
					// Plain strings like "hello \n world" or "first line\n" need to have lineBreaks set to become 2 separate lines as intended
					obj.options.breakLine = true
				}

				// 3: Add CRLF line ending if `breakLine`
				if (obj.options.breakLine && !obj.options.bullet && !obj.options.align && idx + 1 < slideObj.text.length) obj.text += CRLF
			}

			// C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
			if (obj.options.breakLine || obj.text.indexOf(CRLF) > -1) {
				obj.text.split(CRLF).forEach((line, lineIdx) => {
					// Add line-breaks if not bullets/aligned (we add CRLF for those below in STEP 3)
					// NOTE: Use "idx>0" so lines wont start with linebreak (eg:empty first line)
					arrTextObjects.push({
						text: (lineIdx > 0 && obj.options.breakLine && !obj.options.bullet && !obj.options.align ? CRLF : '') + line,
						options: obj.options,
					})
				})
			} else {
				// NOTE: The replace used here is for non-textObjects (plain strings) eg:'hello\nworld'
				arrTextObjects.push(obj)
			}
		})
	}

	// STEP 3: Add bodyProperties
	{
		// A: 'bodyPr'
		strSlideXml += genXmlBodyProperties(slideObj)

		// B: 'lstStyle'
		// NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
		if (opts.h == 0 && opts.line && opts.align) {
			strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>'
		} else if (slideObj.type === 'placeholder') {
			strSlideXml += '<a:lstStyle>'
			strSlideXml += genXmlParagraphProperties(slideObj, true)
			strSlideXml += '</a:lstStyle>'
		} else {
			strSlideXml += '<a:lstStyle/>'
		}
	}

	// STEP 4: Loop over each text object and create paragraph props, text run, etc.
	arrTextObjects.forEach((textObj, idx) => {
		// Clear/Increment loop vars
		paragraphPropXml = '<a:pPr ' + (textObj.options.rtlMode ? ' rtl="1" ' : '')
		textObj.options.lineIdx = idx

		// A: Inherit pPr-type options from parent shape's `options`
		textObj.options.align = textObj.options.align || opts.align
		textObj.options.lineSpacing = textObj.options.lineSpacing || opts.lineSpacing
		textObj.options.indentLevel = textObj.options.indentLevel || opts.indentLevel
		textObj.options.paraSpaceBefore = textObj.options.paraSpaceBefore || opts.paraSpaceBefore
		textObj.options.paraSpaceAfter = textObj.options.paraSpaceAfter || opts.paraSpaceAfter

		textObj.options.lineIdx = idx
		var paragraphPropXml = genXmlParagraphProperties(textObj, false)

		// B: Start paragraph if this is the first text obj, or if current textObj is about to be bulleted or aligned
		if (idx == 0) {
			// Add paragraphProperties right after <p> before textrun(s) begin
			strSlideXml += '<a:p>' + paragraphPropXml
		} else if (idx > 0 && (typeof textObj.options.bullet !== 'undefined' || typeof textObj.options.align !== 'undefined')) {
			strSlideXml += '</a:p><a:p>' + paragraphPropXml
		}

		// C: Inherit any main options (color, fontSize, etc.)
		// We only pass the text.options to genXmlTextRun (not the Slide.options),
		// so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
		Object.entries(opts).forEach(([key, val]) => {
			// NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
			if (key != 'bullet' && !textObj.options[key]) textObj.options[key] = val
		})

		// D: Add formatted textrun
		strSlideXml += genXmlTextRun(textObj)
	})

	// STEP 5: Append 'endParaRPr' (when needed) and close current open paragraph
	// NOTE: (ISSUE#20, ISSUE#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring options
	if (slideObj.type == SLIDE_OBJECT_TYPES.tablecell && (opts.fontSize || opts.fontFace)) {
		if (opts.fontFace) {
			strSlideXml +=
				'<a:endParaRPr lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : '') + ' dirty="0">'
			strSlideXml += '<a:latin typeface="' + opts.fontFace + '" charset="0"/>'
			strSlideXml += '<a:ea typeface="' + opts.fontFace + '" charset="0"/>'
			strSlideXml += '<a:cs typeface="' + opts.fontFace + '" charset="0"/>'
			strSlideXml += '</a:endParaRPr>'
		} else {
			strSlideXml +=
				'<a:endParaRPr lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : '') + ' dirty="0"/>'
		}
	} else {
		strSlideXml += '<a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>' // NOTE: Added 20180101 to address PPT-2007 issues
	}
	strSlideXml += '</a:p>'

	// STEP 6: Close the textBody
	strSlideXml += tagClose

	// LAST: Return XML
	return strSlideXml
}

/**
 * Generate XML Paragraph Properties
 * @param {ISlideObject|IText} textObj - text object
 * @param {boolean} isDefault - array of default relations
 * @return {string} XML
 */
function genXmlParagraphProperties(textObj: ISlideObject | IText, isDefault: boolean): string {
	let strXmlBullet = '',
		strXmlLnSpc = '',
		strXmlParaSpc = ''
	let bulletLvl0Margin = 342900
	let tag = isDefault ? 'a:lvl1pPr' : 'a:pPr'

	let paragraphPropXml = '<' + tag + (textObj.options.rtlMode ? ' rtl="1" ' : '')

	// A: Build paragraphProperties
	{
		// OPTION: align
		if (textObj.options.align) {
			switch (textObj.options.align) {
				case 'left':
					paragraphPropXml += ' algn="l"'
					break
				case 'right':
					paragraphPropXml += ' algn="r"'
					break
				case 'center':
					paragraphPropXml += ' algn="ctr"'
					break
				case 'justify':
					paragraphPropXml += ' algn="just"'
					break
			}
		}

		if (textObj.options.lineSpacing) {
			strXmlLnSpc = '<a:lnSpc><a:spcPts val="' + textObj.options.lineSpacing + '00"/></a:lnSpc>'
		}

		// OPTION: indent
		if (textObj.options.indentLevel && !isNaN(Number(textObj.options.indentLevel)) && textObj.options.indentLevel > 0) {
			paragraphPropXml += ' lvl="' + textObj.options.indentLevel + '"'
		}

		// OPTION: Paragraph Spacing: Before/After
		if (textObj.options.paraSpaceBefore && !isNaN(Number(textObj.options.paraSpaceBefore)) && textObj.options.paraSpaceBefore > 0) {
			strXmlParaSpc += '<a:spcBef><a:spcPts val="' + textObj.options.paraSpaceBefore * 100 + '"/></a:spcBef>'
		}
		if (textObj.options.paraSpaceAfter && !isNaN(Number(textObj.options.paraSpaceAfter)) && textObj.options.paraSpaceAfter > 0) {
			strXmlParaSpc += '<a:spcAft><a:spcPts val="' + textObj.options.paraSpaceAfter * 100 + '"/></a:spcAft>'
		}

		// OPTION: bullet
		// NOTE: OOXML uses the unicode character set for Bullets
		// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
		if (typeof textObj.options.bullet === 'object') {
			if (textObj.options.bullet.type) {
				if (textObj.options.bullet.type.toString().toLowerCase() == 'number') {
					paragraphPropXml +=
						' marL="' +
						(textObj.options.indentLevel && textObj.options.indentLevel > 0
							? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel
							: bulletLvl0Margin) +
						'" indent="-' +
						bulletLvl0Margin +
						'"'
					strXmlBullet = `<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${textObj.options.bullet.style ||
						'arabicPeriod'}" startAt="${textObj.options.bullet.startAt || '1'}"/>`
				}
			} else if (textObj.options.bullet.code) {
				var bulletCode = '&#x' + textObj.options.bullet.code + ';'

				// Check value for hex-ness (s/b 4 char hex)
				if (/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code) == false) {
					console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!')
					bulletCode = BULLET_TYPES['DEFAULT']
				}

				paragraphPropXml +=
					' marL="' +
					(textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel : bulletLvl0Margin) +
					'" indent="-' +
					bulletLvl0Margin +
					'"'
				strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>'
			}
		} else if (textObj.options.bullet == true) {
			paragraphPropXml +=
				' marL="' +
				(textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel : bulletLvl0Margin) +
				'" indent="-' +
				bulletLvl0Margin +
				'"'
			strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + BULLET_TYPES['DEFAULT'] + '"/>'
		} else {
			strXmlBullet = '<a:buNone/>'
		}

		// B: Close Paragraph-Properties
		// IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
		paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet
		if (isDefault) {
			paragraphPropXml += genXmlTextRunProperties(textObj.options, true)
		}
		paragraphPropXml += '</' + tag + '>'
	}

	return paragraphPropXml
}

/**
 * Generate XML Text Run Properties (`a:rPr`)
 * @param {IObjectOptions|ITextOpts} opts - text options
 * @param {boolean} isDefault - whether these are the default text run properties
 * @return {string} XML
 */
function genXmlTextRunProperties(opts: IObjectOptions | ITextOpts, isDefault: boolean): string {
	let runProps = ''
	let runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr'

	// BEGIN runProperties (ex: `<a:rPr lang="en-US" sz="1600" b="1" dirty="0">`)
	runProps += '<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.lang ? ' altLang="en-US"' : '')
	runProps += opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : '' // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
	runProps += opts.bold ? ' b="1"' : ''
	runProps += opts.italic ? ' i="1"' : ''
	runProps += opts.strike ? ' strike="sngStrike"' : ''
	runProps += opts.underline || opts.hyperlink ? ' u="sng"' : ''
	runProps += opts.subscript ? ' baseline="-40000"' : opts.superscript ? ' baseline="30000"' : ''
	runProps += opts.charSpacing ? ' spc="' + opts.charSpacing * 100 + '" kern="0"' : '' // IMPORTANT: Also disable kerning; otherwise text won't actually expand
	runProps += ' dirty="0">'
	// Color / Font / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
	if (opts.color || opts.fontFace || opts.outline) {
		if (opts.outline && typeof opts.outline === 'object') {
			runProps += '<a:ln w="' + Math.round((opts.outline.size || 0.75) * ONEPT) + '">' + genXmlColorSelection(opts.outline.color || 'FFFFFF') + '</a:ln>'
		}
		if (opts.color) runProps += genXmlColorSelection(opts.color)
		if (opts.fontFace) {
			// NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174); ea must come first (Issue #174)
			runProps +=
				'<a:latin typeface="' +
				opts.fontFace +
				'" pitchFamily="34" charset="0"/>' +
				'<a:ea typeface="' +
				opts.fontFace +
				'" pitchFamily="34" charset="-122"/>' +
				'<a:cs typeface="' +
				opts.fontFace +
				'" pitchFamily="34" charset="-120"/>'
		}
	}

	// Hyperlink support
	if (opts.hyperlink) {
		if (typeof opts.hyperlink !== 'object') throw "ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` "
		else if (!opts.hyperlink.url && !opts.hyperlink.slide) throw "ERROR: 'hyperlink requires either `url` or `slide`'"
		else if (opts.hyperlink.url) {
			// FIXME-20170410: FUTURE-FEATURE: color (link is always blue in Keynote and PPT online, so usual text run above isnt honored for links..?)
			//runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
			runProps +=
				'<a:hlinkClick r:id="rId' +
				opts.hyperlink.rId +
				'" invalidUrl="" action="" tgtFrame="" tooltip="' +
				(opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') +
				'" history="1" highlightClick="0" endSnd="0"/>'
		} else if (opts.hyperlink.slide) {
			runProps +=
				'<a:hlinkClick r:id="rId' +
				opts.hyperlink.rId +
				'" action="ppaction://hlinksldjump" tooltip="' +
				(opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') +
				'"/>'
		}
	}

	// END runProperties
	runProps += '</' + runPropsTag + '>'

	return runProps
}

/**
 * Builds `<a:r></a:r>` text runs for `<a:p>` paragraphs in textBody
 * @param {IText} textObj - Text object
 * @return {string} XML string
 */
function genXmlTextRun(textObj: IText): string {
	let arrLines = []
	let paraProp = ''
	let xmlTextRun = ''

	// 1: ADD runProperties
	let startInfo = genXmlTextRunProperties(textObj.options, false)

	// 2: LINE-BREAKS/MULTI-LINE: Split text into multi-p:
	arrLines = textObj.text.split(CRLF)
	if (arrLines.length > 1) {
		arrLines.forEach((line, idx) => {
			xmlTextRun += '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(line)
			// Stop/Start <p>aragraph as long as there is more lines ahead (otherwise its closed at the end of this function)
			if (idx + 1 < arrLines.length) xmlTextRun += (textObj.options.breakLine ? CRLF : '') + '</a:t></a:r>'
		})
	} else {
		// Handle cases where addText `text` was an array of objects - if a text object doesnt contain a '\n' it still need alignment!
		// The first pPr-align is done in makeXml - use line countr to ensure we only add subsequently as needed
		xmlTextRun = (textObj.options.align && textObj.options.lineIdx > 0 ? paraProp : '') + '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(textObj.text)
	}

	// Return paragraph with text run
	return xmlTextRun + '</a:t></a:r>'
}


export function genXmlText(slide: ISlide | ISlideLayout, slideItemObj: ISlideObject,idx:number,placeholderObj: ISlideObject,x,y,cx,cy,locationAttr: string): string {
	let strSlideXml = ''
	let shapeType = null

	//
	if (slideItemObj.shape) shapeType = getShapeInfo(slideItemObj.shape)

	// Lines can have zero cy, but text should not
	if (!slideItemObj.options.line && cy == 0) cy = EMU * 0.3

	// Margin/Padding/Inset for textboxes
	if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
		slideItemObj.options.bodyProp.lIns = slideItemObj.options.margin[0] * ONEPT || 0
		slideItemObj.options.bodyProp.rIns = slideItemObj.options.margin[1] * ONEPT || 0
		slideItemObj.options.bodyProp.bIns = slideItemObj.options.margin[2] * ONEPT || 0
		slideItemObj.options.bodyProp.tIns = slideItemObj.options.margin[3] * ONEPT || 0
	} else if (typeof slideItemObj.options.margin === 'number') {
		slideItemObj.options.bodyProp.lIns = slideItemObj.options.margin * ONEPT
		slideItemObj.options.bodyProp.rIns = slideItemObj.options.margin * ONEPT
		slideItemObj.options.bodyProp.bIns = slideItemObj.options.margin * ONEPT
		slideItemObj.options.bodyProp.tIns = slideItemObj.options.margin * ONEPT
	}

	if (shapeType == null) shapeType = getShapeInfo(null)

	// A: Start SHAPE =======================================================
	strSlideXml += '<p:sp>'

	// B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
	strSlideXml += '<p:nvSpPr><p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '"/>'
	strSlideXml += '<p:cNvSpPr' + (slideItemObj.options && slideItemObj.options.isTextBox ? ' txBox="1"/>' : '/>')
	strSlideXml += '<p:nvPr>'
	strSlideXml += slideItemObj.type === 'placeholder' ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj)
	strSlideXml += '</p:nvPr>'
	strSlideXml += '</p:nvSpPr><p:spPr>'
	strSlideXml += '<a:xfrm' + locationAttr + '>'
	strSlideXml += '<a:off x="' + x + '" y="' + y + '"/>'
	strSlideXml += '<a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
	strSlideXml +=
		'<a:prstGeom prst="' +
		shapeType.name +
		'"><a:avLst>' +
		(slideItemObj.options.rectRadius
			? '<a:gd name="adj" fmla="val ' + Math.round((slideItemObj.options.rectRadius * EMU * 100000) / Math.min(cx, cy)) + '"/>'
			: '') +
		'</a:avLst></a:prstGeom>'

	// Option: FILL
	strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>'

	// shape Type: LINE: line color
	if (slideItemObj.options.line) {
		strSlideXml += '<a:ln' + (slideItemObj.options.lineSize ? ' w="' + slideItemObj.options.lineSize * ONEPT + '"' : '') + '>'
		strSlideXml += genXmlColorSelection(slideItemObj.options.line)
		if (slideItemObj.options.lineDash) strSlideXml += '<a:prstDash val="' + slideItemObj.options.lineDash + '"/>'
		if (slideItemObj.options.lineHead) strSlideXml += '<a:headEnd type="' + slideItemObj.options.lineHead + '"/>'
		if (slideItemObj.options.lineTail) strSlideXml += '<a:tailEnd type="' + slideItemObj.options.lineTail + '"/>'
		strSlideXml += '</a:ln>'
	}

	// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
	if (slideItemObj.options.shadow) {
		slideItemObj.options.shadow.type = slideItemObj.options.shadow.type || 'outer'
		slideItemObj.options.shadow.blur = (slideItemObj.options.shadow.blur || 8) * ONEPT
		slideItemObj.options.shadow.offset = (slideItemObj.options.shadow.offset || 4) * ONEPT
		slideItemObj.options.shadow.angle = (slideItemObj.options.shadow.angle || 270) * 60000
		slideItemObj.options.shadow.color = slideItemObj.options.shadow.color || '000000'
		slideItemObj.options.shadow.opacity = (slideItemObj.options.shadow.opacity || 0.75) * 100000

		strSlideXml += '<a:effectLst>'
		strSlideXml += '<a:' + slideItemObj.options.shadow.type + 'Shdw sx="100000" sy="100000" kx="0" ky="0" '
		strSlideXml += ' algn="bl" rotWithShape="0" blurRad="' + slideItemObj.options.shadow.blur + '" '
		strSlideXml += ' dist="' + slideItemObj.options.shadow.offset + '" dir="' + slideItemObj.options.shadow.angle + '">'
		strSlideXml += '<a:srgbClr val="' + slideItemObj.options.shadow.color + '">'
		strSlideXml += '<a:alpha val="' + slideItemObj.options.shadow.opacity + '"/></a:srgbClr>'
		strSlideXml += '</a:outerShdw>'
		strSlideXml += '</a:effectLst>'
	}

	/* FIXME: FUTURE: Text wrapping (copied from MS-PPTX export)
		// Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
		if ( slideItemObj.options.textWrap ) {
			strSlideXml += '<a:extLst>'
						+ '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
						+ '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1"/>'
						+ '</a:ext>'
						+ '</a:extLst>';
		}
		*/

	// B: Close shape Properties
	strSlideXml += '</p:spPr>'

	// C: Add formatted text (text body "bodyPr")
	strSlideXml += genXmlTextBody(slideItemObj)

	// LAST: Close SHAPE =======================================================
	strSlideXml += '</p:sp>'

	return strSlideXml
}