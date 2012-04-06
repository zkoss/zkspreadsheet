/* Row.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mon Apr 23, 2007 17:29:18 AM , Created by sam
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/

/**
 * 
 * Row represent a row of the spreadsheet, it also be container that contains cells of the row
 */
zss.Row = zk.$extends(zk.Widget, {
	widgetName: 'Row',
	$o: zk.$void, //owner, fellows relationship no needed
	$init: function (sheet, block, row, src) {
		this.$supers(zss.Row, '$init', []); //DO NOT pass "arguments" or all fields will be copied into this Object. 
		
		this.sheet = sheet;
		this.block = block;
		this.src = src;
		this.r = row;
		
		var data = src.getRow(row);
		this.zsh = data.heightId;
		this.cells = [];
		this.wrapedCells = [];
	},
	bind_: function (desktop, skipper, after) {
		this.$supers(zss.Row, 'bind_', arguments);//after bind cells, may need to process wrap height
		var sf = this,
			sheet = this.sheet,
			wgt = sheet._wgt,
			wrapCells = this.wrapedCells;
		if (wrapCells.length) {
			if (wgt.isSheetCSSReady())
				this._updateWrapRowHeight();
			else
				sheet.addSSInitLater(function () {
					sf._updateWrapRowHeight();
				});
			this._listenProcessWrap(true);
		}
		if (zk.ie6_) {
			this.sheet.listen({onRowHeightChanged: this.proxy(this._onRowHeightChanged)});
		}
	},
	unbind_: function () {
		if (zk.ie6_) {
			this.sheet.unlisten({onRowHeightChanged: this.proxy(this._onRowHeightChanged)});
		}
		delete this.cells;
		this.r = this.zsh = null;
		this.$supers(zss.Row, 'unbind_', arguments);
	},
	_listenProcessWrap: function (b) {
		var curr = !!this._processWrap;
		if (curr != b) {
			this.sheet[b ? 'listen' : 'unlisten']({onProcessWrap: this.proxy(this._onProcessWrap)});
			this._processWrap = b;
		}
	},
	_onProcessWrap: function (evt) {
		var d = evt.data,
			r = this.r,
			tRow = d.tRow,
			bRow = d.bRow;
		if (tRow && bRow && r >= tRow && r <= bRow) {
			this._updateWrapRowHeight();
		}
	},
	_updateWrapRowHeight: function () {
		var row = this.r,
			custRowHeight = this.sheet.custRowHeight,
			meta = custRowHeight.getMeta(row),
			orgHgh = custRowHeight._getCustomizedSize(row),
			hgh = orgHgh,
			cells = this.wrapedCells,
			i = cells.length;
		while (i--) {
			var c = cells[i];
			if (c) {
				hgh = Math.max(hgh, c.getTextHeight());
			}
		}
		
//		if (jq(this.$n()).height() == hgh) {
//			return;
//		}
		
		var sheet = this.sheet,
			wgt = sheet._wgt,
			zsh = this.zsh,
			cssId = wgt.getSheetCSSId(),
			pf = wgt.getSelectorPrefix(),
			h2 = (hgh > 0) ? hgh - 1 : 0;
		if (!zsh) {
			zsh = meta ? meta[2] : custRowHeight.ids.next();
			custRowHeight.setCustomizedSize(row, hgh, zsh, false, true);
			sheet._appendZSH(row, zsh); //this doesn't work correctly ?? seems works, TEST it
			sheet._wgt._cacheCtrl.getSelectedSheet().updateRowHeightId(row, zsh);
		} else {
			custRowHeight.setCustomizedSize(row, hgh, zsh, false, true);
		}
			
		zcss.setRule(pf + " .zsh" + zsh, ["height"], [hgh + "px"], true, cssId);
		zcss.setRule(pf + " .zshi" + zsh, "height", h2 + "px", true, cssId);
		zcss.setRule(pf + " .zslh" + zsh, ["display", "height", "line-height"], ["", h2 + "px", h2 + "px"], true, cssId);
		
		//TODO: update focus shall handle by FocusMarkCtrl by listen onRowHeightChanged evt
		var fPos = sheet.getLastFocus(),
			sPos = sheet.getLastSelection();
		if (fPos && sPos) {
			sheet.moveCellFocus(fPos.row, fPos.column, true);
			sheet.moveCellSelection(sPos.left, sPos.top, sPos.right, sPos.bottom, false, true);	
		}
		sheet.fire('onRowHeightChanged', {row: row});
	},
	processWrapCell: function (cell, ignoreUpdateNow) {
		if (this.sheet.custRowHeight.isDefaultSize(cell.r)) {
			var wrapedCells = this.wrapedCells;
			if (cell.wrap) {
				if (!wrapedCells.$contains(cell)) {
					wrapedCells.push(cell);
				}
				if (!ignoreUpdateNow)
					this._updateWrapRowHeight();
			} else {
				if (wrapedCells.$remove(cell)) {
					//if there's no wrap cell to process, shall restore row height by invoke updateWrapRowHeight
					if (!ignoreUpdateNow ||
						!wrapedCells.length) {
						this._updateWrapRowHeight();
					}
				}
			}
			
			var size = wrapedCells.length;
			this._listenProcessWrap(size);
			if (ignoreUpdateNow) //process wrap on sheet onContentChange
				this.sheet.triggerWrap(this.r);
		} 
	},
	//IE6 only
	_onRowHeightChanged: function (evt) {
		if (this.r >= evt.data.row) {
			zk(this.$n()).redoCSS();
		}
	},
	/**
	 * Append zss.Cell at the end of the row
	 * 
	 * @param zss.Cell
	 * @param boolean ignoreDom
	 */
	appendCell: function (cell, ignoreDom) {
		this.appendChild(cell, ignoreDom);
		this.cells.push(cell);
	},
	/**
	 * Insert zss.Cell
	 * 
	 * @param int index
	 * @param zss.Cell
	 * @param boolean ignoreDom
	 */
	insertCell: function (index, cell, ignoreDom) {
		var cells = this.cells,
			sibling = cells[index];
		this.insertBefore(cell, sibling, ignoreDom);
		cells.splice(index, 0, cell);
	},
	redraw: function (out) {
		out.push(this.getHtmlPrologHalf())
		var cells = this.cells,
			size = cells.length;
		for (var i = 0; i < size; i++) {
			var cell = cells[i];
			cell.redraw(out);
		}
		out.push(this.getHtmlEpilogHalf());
	},
	getHtmlPrologHalf: function () {
		return '<div id="' + this.uuid + '" class="' + this.getZclass() + '">';
	},
	getHtmlEpilogHalf: function () {
		return '</div>';
	},
	//super//
	getZclass: function () {
		var cls = 'zsrow',
			hId = this.zsh;
		return hId ? cls + ' zsh' + hId : 'zsrow';
	},
	/**
	 * Returns the {@link zss.Cell}
	 * @param int col column
	 * @return zss.Cell
	 */
	getCell: function (col) {
		var size = this.cells.length,
			i = 0;
		//TODO use binary search
		for (i = 0; i < size; i++) {
			if (this.cells[i].c == col) return this.cells[i];
		}
	},
	/**
	 * Returns the {@link zss.Cell}
	 * @param int index column index
	 * @return zss.Cell
	 */
	getCellAt: function (index) {
		return this.cells[index];
	},
	/**
	 * Remove cell
	 * @param int size
	 */
	removeLeftCell: function (size) {
		var cells = this.cells;
		var beforeSize = cells.length;
		while (size--) {
			if (!cells.length)
				return;
			cells.shift().detach();
		}
	},
	/**
	 * Remove right cell
	 * @param int size
	 */
	removeRightCell: function (size) {
		var cells = this.cells;
		while (size--) {
			if (!cells.length)
				return;
			cells.pop().detach();
		}
	},
	/**
	 * Sets the width position index
	 * @param int index cell index
	 * @param int zsw the width position index
	 */
	appendZSW: function (index, zsw) {
		var cell = this.cells[index];
		cell.appendZSW(zsw);
	},
	/**
	 * Sets the height position index
	 * @param int zsh the height position index
	 */
	appendZSH: function (zsh) {
		if (zsh) {
			this.zsh = zsh;
			jq(this.$n()).addClass("zsh" + zsh);
			var size = this.cells.length;
			for (var i = 0; i < size; i++)
				this.cells[i].appendZSH(zsh);	
		}
	},
	/**
	 * Insert new cell
	 * @param int index cell index
	 * @param int size
	 */
	insertNewCell: function (index, size) {
		var sheet = this.sheet,
			ctrl,
			cells = this.cells,
			col;
		
		if (index > cells.length) return;
		
		//there is a pentional BUG, if index==0 , not template to copy previous format
		//however, for now, the only templat need to copy is overflow and merge cell, 
		//and it is never be care if previous not beend loaded in client.
		var tempcell = index == 0 ? null : cells[index - 1];
		col = index == 0 ? cells[0].c :(tempcell.c + 1);
		
		var block = this.block,
			src = this.src,
			r = this.r;
		
		for (var i = 0; i < size; i++) {
			var c = col + i;
			
			//don't care merge property, it will be sync by removeMergeRange and addMergeRange.
			//don't care the sytle, since cell should be updated by continus updatecell event.
			ctrl = new zss.Cell(sheet, block, r, c, src);
			
			//because of overflow logic, we must maintain overflow link from overhead
			//copy over flow attrbute overby and overhead,
			if (tempcell) {
				if (tempcell.overflowed) ctrl.overlapBy = tempcell;
				else if(tempcell.overlapBy) ctrl.overlapBy = tempcell.overlapBy;
			}
			this.insertCell(index + i, ctrl);
		}
		this.shiftCellInfo(index + size, col + size);
	},
	/**
	 * Shift cell's info
	 * @param int index the start index of the cell
	 * @param int newcol new column index
	 */
	shiftCellInfo: function (index, newcol) {
		var cells = this.cells,
			size = cells.length,
			j = 0;
		for(var i = index; i < size; i++)
			cells[i].resetColumnIndex(newcol+(j++));
	},
	/**
	 * Sets the row index
	 * @param int newrow new row index
	 */
	resetRowIndex: function (newrow) {
		this.r = newrow;
		var cells = this.cells,
			i = cells.length;
		while (i--)
			cells[i].resetRowIndex(newrow);
	},
	/**
	 * Remove the cell of the row
	 * @param int index cell index
	 * @param int size
	 */
	removeCell: function (index, size) {
		var ctrl,
			cells = this.cells,
			col;
		
		if (index > cells.length) return;
		if (index == cells.length)
			col = cells[index - 1].c + 1 ;
		else
			col = cells[index].c;

		var rem = cells.slice(index, index + size),
			tail = cells.slice(index + size, cells.length);
		cells.length = index;
		cells.push.apply(cells, tail);
		
		var cell = rem.pop();
		for (;cell ; cell = rem.pop()) {
			cell.detach();
		}
			
		this.shiftCellInfo(index, col);
	}
}, {
	/**
	 * Returns a row that copy cells from target
	 * @param zss.Row tmprow a target row to copy it's cells
	 * @return string zss.Cell's html content
	 */
	copyCells: function (srcRow, destRow) {
		var cells = srcRow.cells;
			size = cells.length,
			r = destRow.r,
			html = '';
		for (var i = 0; i < size; i++) {
			var srcCell = cells[i],
				c = srcCell.c,
				src = srcCell.src,
				block = srcCell.block,
				sht = srcCell.sheet,
				newCell = new zss.Cell(sht, block, r, c, src);
			html += newCell.getHtml();
			destRow.appendCell(newCell);
		}
		return html;
	}
});
