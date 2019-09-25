
import { genXmlTextBody } from './text'

import { ISlideObject, ISlide, ISlideLayout, ITableCellOpts, ITableCell } from '../core-interfaces'

import {
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_PT,
	EMU,
	ONEPT,
} from '../core-enums'

import { inch2Emu } from '../gen-utils'


let intTableNum: number = 1

export function genXmlTable(slide: ISlide | ISlideLayout, slideItemObj: ISlideObject,x,y,cx,cy): String {
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

	// LAST: Increment counter
	intTableNum++

	return strXml
}