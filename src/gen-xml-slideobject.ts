
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
} from './core-enums'

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
} from './core-interfaces'

import { encodeXmlEntities, inch2Emu, genXmlColorSelection, getSmartParseNumber, convertRotationDegrees } from './gen-utils'

import { PowerPointShapes } from './core-shapes'

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


export function getShapeInfo(shapeName) {
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

/**
 * Generate an XML Placeholder
 * @param {ISlideObject} placeholderObj
 * @returns XML
 */
export function genXmlPlaceholder(placeholderObj: ISlideObject): string {
	if (!placeholderObj) return ''

	let placeholderIdx = placeholderObj.options && placeholderObj.options.placeholderIdx ? placeholderObj.options.placeholderIdx : ''
	let placeholderType = placeholderObj.options && placeholderObj.options.placeholderType ? placeholderObj.options.placeholderType : ''

	return `<p:ph
		${placeholderIdx ? ' idx="' + placeholderIdx + '"' : ''}
		${placeholderType && PLACEHOLDER_TYPES[placeholderType] ? ' type="' + PLACEHOLDER_TYPES[placeholderType] + '"' : ''}
		${placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : ''}
		/>`
}

export function serialiseSlideObject(slide: ISlide | ISlideLayout, slideItemObj: ISlideObject, idx: number): String {
	let x = 0,
		y = 0,
		cx = getSmartParseNumber('75%', 'X', slide.presLayout),
		cy = 0
	let placeholderObj: ISlideObject
	let locationAttr = ''
	let shapeType = null
	let strSlideXml = ''

	// FIXME THIS IS WRONG!!
	let intTableNum: number = 1


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
	if (slideItemObj.shape) shapeType = getShapeInfo(slideItemObj.shape)
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
			slideItemObj.data.map( serialiseSlideObject.bind(null,slide) ).forEach( xml => strSlideXml += xml )
			strSlideXml += '</p:grpSp>'
			break;
		case SLIDE_OBJECT_TYPES.table:
			let objTableGrid = {}
			let arrTabRows = slideItemObj.arrTabRows
			let objTabOpts = slideItemObj.options
			let intColCnt = 0,
				intColW = 0
			let cellOpts: ITableCellOpts

			// Calc number of columns
			// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
			// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
			arrTabRows[0].forEach(cell => {
				cellOpts = cell.options || null
				intColCnt += cellOpts && cellOpts.colspan ? Number(cellOpts.colspan) : 1
			})

			// STEP 1: Start Table XML
			// NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
			let strXml =
				'<p:graphicFrame>' +
				'  <p:nvGraphicFramePr>' +
				'    <p:cNvPr id="' +
				(intTableNum * slide.number + 1) +
				'" name="Table ' +
				intTableNum * slide.number +
				'"/>' +
				'    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
				'    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>' +
				'  </p:nvGraphicFramePr>' +
				'  <p:xfrm>' +
				'    <a:off x="' +
				(x || (x == 0 ? 0 : EMU)) +
				'" y="' +
				(y || (y == 0 ? 0 : EMU)) +
				'"/>' +
				'    <a:ext cx="' +
				(cx || (cx == 0 ? 0 : EMU)) +
				'" cy="' +
				(cy || EMU) +
				'"/>' +
				'  </p:xfrm>' +
				'  <a:graphic>' +
				'    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
				'      <a:tbl>' +
				'        <a:tblPr/>'
			// + '        <a:tblPr bandRow="1"/>';
			// FIXME: Support banded rows, first/last row, etc.
			// NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
			// <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">

			// STEP 2: Set column widths
			// Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
			// A: Col widths provided?
			if (Array.isArray(objTabOpts.colW)) {
				strXml += '<a:tblGrid>'
				for (var col = 0; col < intColCnt; col++) {
					strXml +=
						'<a:gridCol w="' +
						Math.round(inch2Emu(objTabOpts.colW[col]) || (typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt) +
						'"/>'
				}
				strXml += '</a:tblGrid>'
			}
			// B: Table Width provided without colW? Then distribute cols
			else {
				intColW = objTabOpts.colW ? objTabOpts.colW : EMU
				if (slideItemObj.options.w && !objTabOpts.colW) intColW = Math.round((typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt)
				strXml += '<a:tblGrid>'
				for (var col = 0; col < intColCnt; col++) {
					strXml += '<a:gridCol w="' + intColW + '"/>'
				}
				strXml += '</a:tblGrid>'
			}

			// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
			// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
			// so a simple loop below in XML building wont suffice to build table correctly.
			// We have to build an actual grid now
			/*
				EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

				/------|------|------|------\
				|  A0  |  B0  |  C0  |  D0  |
				|      |  B1  |  C1  |      |
				|      |      |  C2  |  D2  |
				\------|------|------|------/
			*/
			/*
				Object ex: key = rowIdx / val = [cells] cellIdx { 0:{type: "tablecell", text: Array(1), options: {…}}, 1:... }
				{0: {…}, 1: {…}, 2: {…}, 3: {…}}
			*/
			arrTabRows.forEach((row, rIdx) => {
				// A: Create row if needed (recall one may be created in loop below for rowspans, so dont assume we need to create one each iteration)
				if (!objTableGrid[rIdx]) objTableGrid[rIdx] = {}

				// B: Loop over all cells
				row.forEach((cell, cIdx) => {
					// DESIGN: NOTE: Row cell arrays can be "uneven" (diff cell count in each) due to rowspan/colspan
					// Therefore, for each cell we run 0->colCount to determine the correct slot for it to reside
					// as the uneven/mixed nature of the data means we cannot use the cIdx value alone.
					// E.g.: the 2nd element in the row array may actually go into the 5th table grid row cell b/c of colspans!
					for (var idx = 0; cIdx + idx < intColCnt; idx++) {
						var currColIdx = cIdx + idx

						if (!objTableGrid[rIdx][currColIdx]) {
							// A: Set this cell
							objTableGrid[rIdx][currColIdx] = cell

							// B: Handle `colspan` or `rowspan` (a {cell} cant have both! FIXME: FUTURE: ROWSPAN & COLSPAN in same cell)
							if (cell && cell.options && cell.options.colspan && !isNaN(Number(cell.options.colspan))) {
								for (var idy = 1; idy < Number(cell.options.colspan); idy++) {
									objTableGrid[rIdx][currColIdx + idy] = { hmerge: true, text: 'hmerge' }
								}
							} else if (cell && cell.options && cell.options.rowspan && !isNaN(Number(cell.options.rowspan))) {
								for (var idz = 1; idz < Number(cell.options.rowspan); idz++) {
									if (!objTableGrid[rIdx + idz]) objTableGrid[rIdx + idz] = {}
									objTableGrid[rIdx + idz][currColIdx] = { vmerge: true, text: 'vmerge' }
								}
							}

							// C: Break out of colCnt loop now that slot has been filled
							break
						}
					}
				})
			})

			/* DEBUG: useful for rowspan/colspan testing
			if ( objTabOpts.verbose ) {
				console.table(objTableGrid);
				var arrText = [];
				objTableGrid.forEach(function(row){ let arrRow = []; row.forEach(row,function(cell){ arrRow.push(cell.text); }); arrText.push(arrRow); });
				console.table( arrText );
			}
			*/

			// STEP 4: Build table rows/cells
			Object.entries(objTableGrid).forEach(([rIdx,rowObj]) => {
				// A: Table Height provided without rowH? Then distribute rows
				let intRowH = 0 // IMPORTANT: Default must be zero for auto-sizing to work
				if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]))
				else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH))) intRowH = inch2Emu(Number(objTabOpts.rowH))
				else if (slideItemObj.options.cy || slideItemObj.options.h)
					intRowH =
						(slideItemObj.options.h ? inch2Emu(slideItemObj.options.h) : typeof slideItemObj.options.cy === 'number' ? slideItemObj.options.cy : 1) /
						arrTabRows.length

				// B: Start row
				strXml += '<a:tr h="' + intRowH + '">'

				// C: Loop over each CELL
				Object.entries(rowObj).forEach(([_cIdx,cellObj]) => {
					let cell:ITableCell = cellObj

					// 1: "hmerge" cells are just place-holders in the table grid - skip those and go to next cell
					if (cell.hmerge) return

					// 2: OPTIONS: Build/set cell options
					let cellOpts = cell.options || ({} as ITableCell['options'])
					/// TODO-3: FIXME: ONLY MAKE CELLS with objects! if (typeof cell === 'number' || typeof cell === 'string') cell = { text: cell.toString() }
					cell.options = cellOpts

					// B: Inherit some options from table when cell options dont exist
					// @see: http://officeopenxml.com/drwTableCellProperties-alignment.php
					;['align', 'bold', 'border', 'color', 'fill', 'fontFace', 'fontSize', 'margin', 'underline', 'valign'].forEach(name => {
						if (objTabOpts[name] && !cellOpts[name] && cellOpts[name] != 0) cellOpts[name] = objTabOpts[name]
					})

					let cellValign = cellOpts.valign
						? ' anchor="' +
						  cellOpts.valign
								.replace(/^c$/i, 'ctr')
								.replace(/^m$/i, 'ctr')
								.replace('center', 'ctr')
								.replace('middle', 'ctr')
								.replace('top', 't')
								.replace('btm', 'b')
								.replace('bottom', 'b') +
						  '"'
						: ''
					let cellColspan = cellOpts.colspan ? ' gridSpan="' + cellOpts.colspan + '"' : ''
					let cellRowspan = cellOpts.rowspan ? ' rowSpan="' + cellOpts.rowspan + '"' : ''
					let cellFill =
						(cell.optImp && cell.optImp.fill) || cellOpts.fill
							? ' <a:solidFill><a:srgbClr val="' +
							  ((cell.optImp && cell.optImp.fill) || (typeof cellOpts.fill === 'string' ? cellOpts.fill.replace('#', '') : '')).toUpperCase() +
							  '"/></a:solidFill>'
							: ''
					let cellMargin = cellOpts.margin == 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_PT
					if (!Array.isArray(cellMargin) && typeof cellMargin === 'number') cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin]
					let cellMarginXml =
						' marL="' +
						cellMargin[3] * ONEPT +
						'" marR="' +
						cellMargin[1] * ONEPT +
						'" marT="' +
						cellMargin[0] * ONEPT +
						'" marB="' +
						cellMargin[2] * ONEPT +
						'"'

					// FIXME: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatever options exist)

					// 3: ROWSPAN: Add dummy cells for any active rowspan
					if (cell.vmerge) {
						strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>'
						return
					}

					// 4: Set CELL content and properties ==================================
					strXml += '<a:tc' + cellColspan + cellRowspan + '>' + genXmlTextBody(cell) + '<a:tcPr' + cellMarginXml + cellValign + '>'

					// 5: Borders: Add any borders
					/// TODO=3: FIXME: stop using `none` if (cellOpts.border && typeof cellOpts.border === 'string' && cellOpts.border.toLowerCase() == 'none') {
					if (cellOpts.border && !Array.isArray(cellOpts.border) && cellOpts.border.type == 'none') {
						strXml += '  <a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>'
						strXml += '  <a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR>'
						strXml += '  <a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT>'
						strXml += '  <a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>'
					} else if (cellOpts.border && typeof cellOpts.border === 'string') {
						strXml +=
							'  <a:lnL w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnL>'
						strXml +=
							'  <a:lnR w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnR>'
						strXml +=
							'  <a:lnT w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnT>'
						strXml +=
							'  <a:lnB w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnB>'
					} else if (cellOpts.border && Array.isArray(cellOpts.border)) {
						[{ idx: 3, name: 'lnL' }, { idx: 1, name: 'lnR' }, { idx: 0, name: 'lnT' }, { idx: 2, name: 'lnB' }].forEach(obj => {
							if (cellOpts.border[obj.idx]) {
								let strC =
									'<a:solidFill><a:srgbClr val="' +
									(cellOpts.border[obj.idx].color ? cellOpts.border[obj.idx].color : DEF_CELL_BORDER.color) +
									'"/></a:solidFill>'
								let intW =
									cellOpts.border[obj.idx] && (cellOpts.border[obj.idx].pt || cellOpts.border[obj.idx].pt == 0)
										? ONEPT * Number(cellOpts.border[obj.idx].pt)
										: ONEPT
								strXml += '<a:' + obj.name + ' w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strC + '</a:' + obj.name + '>'
							} else strXml += '<a:' + obj.name + ' w="0"><a:miter lim="400000"/></a:' + obj.name + '>'
						})
					} else if (cellOpts.border && !Array.isArray(cellOpts.border)) {
						let intW = cellOpts.border && (cellOpts.border.pt || cellOpts.border.pt == 0) ? ONEPT * Number(cellOpts.border.pt) : ONEPT
						let strClr =
							'<a:solidFill><a:srgbClr val="' +
							(cellOpts.border.color ? cellOpts.border.color.replace('#', '') : DEF_CELL_BORDER.color) +
							'"/></a:solidFill>'
						let strAttr = '<a:prstDash val="'
						strAttr += cellOpts.border.type && cellOpts.border.type.toLowerCase().indexOf('dash') > -1 ? 'sysDash' : 'solid'
						strAttr += '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>'
						// *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
						strXml += '<a:lnL w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnL>'
						strXml += '<a:lnR w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnR>'
						strXml += '<a:lnT w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnT>'
						strXml += '<a:lnB w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnB>'
						// *** IMPORTANT! *** LRTB order matters!
					}

					// 6: Close cell Properties & Cell
					strXml += cellFill
					strXml += '  </a:tcPr>'
					strXml += ' </a:tc>'

					// LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
					if (cellOpts.colspan) {
						for (var tmp = 1; tmp < Number(cellOpts.colspan); tmp++) {
							strXml += '<a:tc hMerge="1"><a:tcPr/></a:tc>'
						}
					}
				})

				// D: Complete row
				strXml += '</a:tr>'
			})

			// STEP 5: Complete table
			strXml += '      </a:tbl>'
			strXml += '    </a:graphicData>'
			strXml += '  </a:graphic>'
			strXml += '</p:graphicFrame>'

			// STEP 6: Set table XML
			strSlideXml += strXml

			// LAST: Increment counter
			intTableNum++
			break

		case SLIDE_OBJECT_TYPES.text:
		case SLIDE_OBJECT_TYPES.placeholder:
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
			break

		case SLIDE_OBJECT_TYPES.image:
			var sizing = slideItemObj.options.sizing,
				rounding = slideItemObj.options.rounding,
				width = cx,
				height = cy

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
			break

		case SLIDE_OBJECT_TYPES.media:
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
			break

		case SLIDE_OBJECT_TYPES.chart:
			strSlideXml += '<p:graphicFrame>'
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
			break
	}

	return strSlideXml
}