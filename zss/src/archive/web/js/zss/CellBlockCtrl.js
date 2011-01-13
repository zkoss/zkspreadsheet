/* CellBlockCtrl.js

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
 * CellBlockCtrl is used for controlling cells, include load cell, creating cell, merge cell and cell position index
 */
zss.CellBlockCtrl = zk.$extends(zk.Object, {
	range: null,
	$init: function (sheet, cmp, left, top) {
		this.$supers('$init', arguments);
		this.sheet = sheet;
		this.sheetid = sheet.sheetid;
		this.comp = cmp;
		zk(cmp).disableSelection();//disable block is selectable

		this.range = new zss.Range(left, top, 0, 0, true);
		this._overflowcell = [];//cell is overflow
		this._textedcell = [];//cell has text
		this._lastDir = null;
		this.rows = [];
		cmp.ctrl = this;
	},
	cleanup: function (cmp) {
		this.invalid = true;
		this.sheet = null;	
		if (this.comp) this.comp.ctrl = null;
		this.comp = null;
		
		var i = this.rows.length;
		while (i--) {
			this.rows[i].cleanup();
			this.rows[i] = null;
		}
		
		//this.rowcmps = null;
		this.rows.length = this._overflowcell.length = this._textedcell.length = 0;
		this.rows = this._overflowcell = this._textedcell = null;
	},
	/**
	 * Returns the cell of the spreadsheet
	 * @param int row
	 * @param int col
	 */
	getCell: function (row, col) {
		var range = this.range;
		if(row < range.top || row > range.bottom || col < range.left || col > range.right)
			return null;
		
		return this.rows[row - range.top].getCellAt(col - range.left);
	},
	/**
	 * Sets the merge range of the cell
	 * @param id
	 * @param int left column start index
	 * @param int top row start index
	 * @param int right column end index
	 * @param int bottom row end index
	 * @param int width
	 */
	addMergeRange: function (id, left, top, right, bottom, width) {
		var cell = cell1 = this.getCell(top, left);
		if (!cell) return;
		var comp;
		for (var i = left; i <= right; i++) {
			if (i == left) {
				comp = cell.comp;
				jq(comp).addClass("zsmerge" + id);
			} else {
				cell = this.getCell(top, i);
				if (!cell) break;//in top / left /corner , they might not have such cell
				comp = cell.comp;
				jq(comp).addClass("zsmergee");
			}
			jq(comp).attr({"z.merid": id, "z.merr": right, "z.merl": left});

			cell.merid = id;
			cell.merr = right;
			cell.merl = left;
		}
		zss.Cell._processOverflow(cell1);
	},
	/**
	 * Remove the merge range of the cell
	 * @param id
	 * @param int left column start index
	 * @param int top row start index
	 * @param int right column end index
	 * @param int bottom row end index
	 * @param int width
	 */
	removeMergeRange: function (id, left, top, right, bottom, width) {
		var cell = cell1 = this.getCell(top, left);
		if (!cell) return;
		var merl = cell.merl,
			merr = cell.merr,
			merid = cell.merid;
		
		if (id != merid)
			return;

		var ud,
			comp;
		for (var i = merl; i <= merr; i++) {
			if (i == merl) {
				comp = cell.comp;
				jq(comp).removeClass("zsmerge" + id);
			} else {
				cell = this.getCell(top, i);
				if (!cell) break;//in top / left /corner , they might not have such cell
				comp = cell.comp;
				jq(comp).removeClass("zsmergee");
			}
			jq(comp).removeAttr("z.merid").removeAttr("z.merr").removeAttr("z.merl");
			cell.merid = cell.merr = cell.merl = ud;
		}
		zss.Cell._processOverflow(cell1);
	},
	_processTextOverflow: function (colbefore) {
		var sheet = this.sheet;
		if (!sheet.config.textOverflow) return;
		
		var txtcells = this._textedcell,
			size = txtcells.length,
			i = 0;
		for (; i < size; i++) {
			var ctrl = txtcells[i];
			if(/*!ctrl.wrap && */(!colbefore || ctrl.c <= colbefore)){
				//TODO i shoud process the closer cell only.
				zss.Cell._processOverflow(ctrl);
			}
		}
	},	
	_createEastCell: function (blockdata, width, height) {
		for (var i = 0 ; i < height; i++) {
			var type = blockdata[i].type,
				rowindex = blockdata[i].ix,
				cells = blockdata[i].cells,
				rowcmp = this.rows[i].comp;
			
			jq(rowcmp).css('display', 'none');//for speed up
			for (var j = 0; j < width; j++) {
				var cell = cells[j],
					parm ={row: rowindex, col: cell.ix, zsh: blockdata[i].zsh};
				zkS.copyParm(cell, parm, ["txt", "st", "ist", "wrap", "hal", "rbo", "merr", "merid", "merl", "zsw", "edit"]);
				
				var cellctrl = zss.Cell.createComp(this.sheet, this, parm);
				rowcmp.ctrl.pushCellE(cellctrl);
			}
			jq(rowcmp).css('display', '');
		}
		this.range.extendRight(width);
	},
	_createWestCell: function (blockdata, width, height) {
		for(var i = 0; i < height; i++) {
			var type = blockdata[i].type,
				rowindex = blockdata[i].ix,
				cells = blockdata[i].cells;
			
			var rowcmp = this.rows[i].comp;
			jq(rowcmp).css('display', 'none');//for speed up
			for (var j = 0; j < width; j++) {
				var cell = cells[j],
					parm ={row: rowindex, col: cell.ix, zsh: blockdata[i].zsh};
				zkS.copyParm(cell, parm, ["txt", "st", "ist", "wrap", "hal", "rbo", "merr", "merid", "merl", "zsw", "edit"]);
				
				var cellctrl = zss.Cell.createComp(this.sheet, this, parm);
				rowcmp.ctrl.pushCellS(cellctrl);
			}
			jq(rowcmp).css('display', '');
		}
		this.range.extendLeft(width);
	},
	_createNorthCell: function (blockdata, width, height) {
		var rowctrls = [];

		for (var i = 0; i < height; i++) {
			var type = blockdata[i].type,
				rowindex = blockdata[i].ix,
				cells = blockdata[i].cells,
				parm = {row: rowindex};
			zkS.copyParm(blockdata[i], parm, ["zsh"]);
			var ctrl = new zss.Row.createComp(this.sheet, this,parm),
				rowcmp =  ctrl.comp;
			//this.rowcmps[rowindex] = rowcmp;
			for (var j = 0; j < width; j++) {
				var cell = cells[j];
				parm = {row: rowindex, col: cell.ix, zsh: blockdata[i].zsh};
				zkS.copyParm(cell, parm, ["txt", "st", "ist", "wrap", "hal", "rbo", "merr", "merid", "merl", "zsw", "edit"]);

				var cellctrl = zss.Cell.createComp(this.sheet, this,parm);
				rowcmp.ctrl.pushCellE(cellctrl);	
			}
			rowctrls.push(ctrl);
		}
		for (var i = 0; i < height; i++)
			this.pushRowS(rowctrls[i]);

		this.range.extendTop(height);
	},
	_createSouthCell: function (blockdata, width, height) {
		var rowctrls = [];
		for (var i = 0; i < height; i++) {
			var type = blockdata[i].type,
				rowindex = blockdata[i].ix,
				cells = blockdata[i].cells,
				parm = {row:rowindex};
			zkS.copyParm(blockdata[i], parm, ["zsh"]);
			
			var ctrl = new zss.Row.createComp(this.sheet, this, parm),
				rowcmp =  ctrl.comp;
			//this.rowcmps[rowindex] = rowcmp;
			for (var j = 0; j < width; j++) {
				var cell = cells[j],
					parm = {row: rowindex, col: cell.ix, zsh: blockdata[i].zsh};
				zkS.copyParm(cell, parm, ["txt", "st", "ist", "wrap", "hal", "rbo", "merr", "merid", "merl", "zsw", "edit"]);
				
				var cellctrl = zss.Cell.createComp(this.sheet, this, parm);
				rowcmp.ctrl.pushCellE(cellctrl);	
			}
			rowctrls.push(ctrl);
		}
		
		for (var i = 0; i < height; i++)
			this.pushRowE(rowctrls[i]);
		
		this.range.extendBottom(height);
	},
	_createJumpCell : function (blockdata, width, height) {
		var rowctrls = [];
		for (var i = 0; i < height; i++) {
			var type = blockdata[i].type,
				rowindex = blockdata[i].ix,
				cells = blockdata[i].cells,
				parm ={row: rowindex};
			zkS.copyParm(blockdata[i], parm, ["zsh"]);
			
			var ctrl = new zss.Row.createComp(this.sheet, this, parm),
				rowcmp =  ctrl.comp;
			//this.rowcmps[rowindex] = rowcmp;
			for (var j = 0; j < width; j++) {
				var cell = cells[j];
					parm = {row: rowindex, col: cell.ix, zsh:blockdata[i].zsh};
				zkS.copyParm(cell, parm, ["txt", "st", "ist", "wrap", "hal", "rbo", "merr", "merid", "merl", "zsw", "edit"]);
				
				var cellctrl = zss.Cell.createComp(this.sheet, this, parm);
				rowcmp.ctrl.pushCellE(cellctrl);	
			}
			rowctrls.push(ctrl);
		}
		for (var i = 0; i < height; i++) {
			this.pushRowE(rowctrls[i]);
		}
		var oldRange = this.range;
		this.range = new zss.Range(oldRange.left, oldRange.top, width, height, true);
	},
	/**
	 * Sets rows's width position index
	 * @param int col the column
	 * @param int zsw the column's position index
	 */
	appendZSW: function (col, zsw) {
		if (col > this.range.right || col < this.range.left) return;
		rows = this.rows;
		var rowsize = rows.length;
		for (var i = 0; i < rowsize; i++)
			rows[i].appendZSW(col - this.range.left, zsw);
	},
	/**
	 * Sets rows's height position index
	 * @param int row the row
	 * @param int zsh the row's height position index
	 */
	appendZSH: function (row, zsh) {
		if (row > this.range.bottom || row < this.range.top) return;
		this.rows[row - this.range.top].appendZSH(zsh);
	},
	/**
	 * Insert new column
	 * @param int col the column to insert
	 * @param int size the size of the cell
	 */
	insertNewColumn: function (col, size) {
		if (col > (this.range.right + 1) || col < this.range.left) return;
		var index = col - this.range.left;
		
		rows = this.rows;
		var rowsize = rows.length;
		for (var i = 0; i < rowsize; i++)
			rows[i].insertNewCell(index, size);

		this.range.extendRight(size);
	},
	/**
	 * Insert new row
	 * @param int row the row to insert
	 * @param int size the size of the row  
	 */
	insertNewRow: function (row, size) {
		if (row > (this.range.bottom + 1) || row < this.range.top) return;
		var index = row - this.range.top,
			ctrl,
			rows = this.rows,
			temprow = (index >= rows.length) ? rows[rows.length - 1] : rows[index],
			sheet = this.sheet,
			parm = {};
		for (var i = 0; i < size; i++) {
			parm.row = row + i;
			ctrl = zss.Row.createComp(this.sheet, this, parm);
			zss.Row.copyCells(temprow, ctrl);
			this.pushRowI(ctrl, index + i);
		}
		this.shiftRowInfo(index + size, row + size);
		this.range.extendBottom(size);
	},
	/**
	 * Shift row info
	 */
	shiftRowInfo: function(index, newrow) {
		var rows = this.rows,
			size = this.rows.length,
			j = 0;
		for(var i = index; i < size; i++)
			rows[i].resetRowIndex(newrow +(j++));
	},
	/**
	 * Remove column
	 * @param int col the column to remove
	 * @param int size
	 */
	removeColumn: function (col, size) {
		if (col > (this.range.right + 1) || col < this.range.left) return;
		var index = col - this.range.left;
		if ((col + size) > this.range.right)
			size = this.range.right - col + 1;
		
		rows = this.rows;
		var rowsize = rows.length;
		for (var i = 0; i < rowsize; i++)
			rows[i].removeCell(index, size);
		this.range.extendRight(-size);
	},
	/**
	 * Remove row
	 * @param int row the row to remove
	 * @param int size
	 */
	removeRow: function (row, size) {
		if (row > (this.range.bottom + 1) || row < this.range.top) return;
		var index = row - this.range.top;
		if ((row + size) > this.range.bottom)
			size = this.range.bottom - row + 1;
		
		var ctrl,
			rows = this.rows,
			rem = rows.slice(index, index + size),
			tail = rows.slice(index + size, rows.length);
		rows.length = index;
		rows.push.apply(rows, tail);

		var ctrl = rem.pop();

		for (; ctrl; ctrl = rem.pop()) {
			jq(ctrl.comp).remove();
			ctrl.cleanup();
		}
		this.range.extendBottom(-size);
		this.shiftRowInfo(index, row);
	},
	/**
	 * Sets the row to the end of the header
	 * @param zss.Row row
	 */
	pushRowE: function (row) {
		this.rows.push(row);
		this.comp.appendChild(row.comp);
	},
	/**
	 * Sets the row to the start of the header
	 * @param zss.Row row
	 */
	pushRowS: function (row) {
		this.rows.unshift(row);
		this.comp.insertBefore(row.comp, this.comp.firstChild);
	},
	/**
	 * Sets the row by index
	 * @param zss.Row
	 * @param int index
	 */
	/** insert cell, TODO cellctrls for better performance **/
	pushRowI: function (rowctrl, index) {
		var rows = this.rows,
			size = rows.length;
		if (index > size)
			throw('index out of bound:' + index + ' > ' + size );

		if (index == 0)
			this.pushRowS(rowctrl);
		else if (index == size)
			this.pushRowE(rowctrl);
		else {
			var tail = rows.slice(index, size);
			rows.length = index;
			rows.push(rowctrl);
			rows.push.apply(rows, tail);
			this.comp.insertBefore(rowctrl.comp, tail[0].comp);	
		}
	},
	/** load existed for element, this method is called when initial.**/
	loadByComp: function (cmp) {
		
		this.comp = cmp;
		cmp.ctrl = this;
		var next = jq(cmp).children("DIV:first")[0],
			minx = miny = maxx = maxy = 0;
		while (next) {
			if (jq(next).attr('zs.t') == "SRow") {
				var ctrl = new zss.Row(this.sheet, this, next);
				this.rows.push(ctrl);
				var r = zk.parseInt(jq(next).attr("z.r")),
					nextcell = jq(next).children('DIV:first')[0],
					editrow = this.sheet._wgt._edittext[r];
				while (nextcell) {
					var c = zk.parseInt(jq(nextcell).attr("z.c"));
					if (minx > c)
						minx = c;
					if (maxx < c)
						maxx = c;
					var cell = new zss.Cell(this.sheet, this, nextcell, editrow[c]);
					ctrl.pushCell(cell);
					nextcell = jq(nextcell).next('DIV')[0];
				}
				
				if (miny > r)
					miny = r;
				if (maxy < r)
					maxy = r;

				zkS.trimFirst(next, "DIV");
				zkS.trimLast(next, "DIV");
				
				next = jq(next).next('DIV')[0];
				continue;	
			}
		}
		var cleft = minx,
			ctop = miny,
			cw = maxx - minx + 1,
			ch = maxy - miny + 1;
		this.range = new zss.Range(cleft, ctop, cw, ch, true);
	},
	/**
	 * Hide the cell
	 */
	hide: function () {
		jq(this.comp).css('display', 'none');
	},
	/**
	 * Show the cell
	 */
	show : function () {
		jq(this.comp).css('display', 'block');
	},
	_removeWestCell: function (size) {
		rows = this.rows;
		var rowsize = this.rows.length,
			i = 0;
		for(; i < rowsize; i++)
			this.rows[i].removeLeftCell(size);
		this.range.extendLeft(-size);
	},
	_removeEastCell: function (size) {
		rows = this.rows;
		var rowsize = this.rows.length,
			i = 0;
		for(; i < rowsize; i++)
			this.rows[i].removeRightCell(size);
		this.range.extendRight(-size);
	},
	_removeNorthCell: function (size) {
		rows = this.rows;
		var cmp = this.comp,
			rowsize = this.rows.length,
			i = 0,
			rowctrl;
		for (; i < size; i++) {
			rowctrl = rows.shift();
			jq(rowctrl.comp).remove();
			rowctrl.cleanup();
		}
		this.range.extendTop(-size);
	},
	_removeSouthCell: function (size) {
		rows = this.rows;
		var cmp = this.comp,
			rowsize = this.rows.length,
			i = 0,
			rowctrl;
		for(; i < size; i++) {
			rowctrl = rows.pop();
			jq(rowctrl.comp).remove();
			rowctrl.cleanup();
		}
		this.range.extendBottom(-size);
	}
}, {
	createComp: function (sheet, left, top, sclass) {
	/*
	<div zs.t="SBlock" class="zsblock">
	*/	
		var cmp = document.createElement("div");
		jq(cmp).attr("zs.t", "SBlock").addClass(!sclass ? "zsblock" : sclass);
		return new zss.CellBlockCtrl(sheet, cmp, left, top);
	}
});