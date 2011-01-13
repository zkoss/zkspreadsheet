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
zss.Row = zk.$extends(zk.Object, {
	$init: function (sheet, block, cmp) {
		this.$supers('$init', arguments);
		this.sheet = sheet;
		this.sheetid = sheet.id;
		this.block = block;
		this.comp= cmp;
		this.r = zk.parseInt(jq(cmp).attr("z.r"));
		this.zsh = jq(cmp).attr("z.zsh");
		if (this.zsh)
			this.zsh = zk.parseInt(this.zsh);
		this.cells = [];
		cmp.ctrl = this;
	},
	cleanup: function (){
		this.invalid = true;

		if (this.comp) this.comp.ctrl = null;
		this.comp = this.sheet = this.block = null;

		var size = this.cells.length,
			i = 0;
		for (i = 0; i < size; i++) {
			this.cells[i].cleanup();
			this.cells[i] = null;
		}
		this.cells.length = 0;
		this.cells = null;
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
	 * @param int column index
	 * @return zss.Cell
	 */
	getCellAt: function (index) {
		return this.cells[index];
	},
	/** 
	 * Push cell to the end, don't append component
	 * @param zss.Cell cell
	 */
	pushCell: function (cell) {
		this.cells.push(cell);
	},
	/**
	 * Push cell/cells to the end, append component
	 * @param zss.Cell/array cell
	 */
	pushCellE: function (cell) {
		this.cells.push(cell);
		this.comp.appendChild(cell.comp);
	},
	/**
	 * Push cell to the start, append component
	 * @param zss.Cell
	 */
	pushCellS: function (cell) {
		this.cells.unshift(cell);
		this.comp.insertBefore(cell.comp, this.comp.firstChild);
	},
	/**
	 * Insert cell, TODO cellctrls for better performance
	 * @param zss.Cell cellctrl
	 * @param int index
	 */
	pushCellI: function (cellctrl, index) {
		var cells = this.cells,
			size = cells.length;
		if (index > size)
			throw('index out of bound:' + index + ' > ' + size) ;

		if (index == 0)
			this.pushCellS(cellctrl);
		else if (index == size)
			this.pushCellE(cellctrl);
		else {
			var tail = cells.slice(index,size);
			cells.length = index;
			cells.push(cellctrl);
			cells.push.apply(cells, tail);
			this.comp.insertBefore(cellctrl.comp, tail[0].comp);	
		}
	},
	/**
	 * Remove left cell
	 * @param int size
	 */
	removeLeftCell: function (size) {
		var cells = this.cells,
			cell;
		while (size--) {
			cell = cells.shift();
			jq(cell.comp).remove();
			cell.cleanup();
		}
	},
	/**
	 * Remove right cell
	 * @param int size
	 */
	removeRightCell: function (size) {
		var cells = this.cells,
			cell;
		while (size--) {
			cell = cells.pop();
			jq(cell.comp).remove();
			cell.cleanup();
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
		this.zsh = zsh;
		jq(this.comp).addClass("zsh" + zsh).attr("z.zsh", zsh);
		var size = this.cells.length;
		for (var i = 0; i < size; i++)
			this.cells[i].appendZSH(zsh);
	},
	/**
	 * Insert new cell
	 * @param int index cell index
	 * @param int size
	 */
	insertNewCell: function (index, size) {
		var ctrl,
			cells = this.cells,
			col;
		
		if (index > cells.length) return;
		
		//there is a pentional BUG, if index==0 , not template to copy previous format
		//however, for now, the only templat need to copy is overflow and merge cell, 
		//and it is never be care if previous not beend loaded in client.
		var tempcell = index ==0 ? null : cells[index - 1];
		col = index == 0 ? cells[0].c :(tempcell.c + 1);
		
		var parm = {row: this.r};
		if (this.zsh)
			parm.zsh = this.zsh;
		
		for (var i = 0; i < size; i++) {
			parm.col = col + i;
			//don't care merge property, it will be sync by removeMergeRange and addMergeRange.
			//don't care the sytle, since cell should be updated by continus updatecell event.
			ctrl = zss.Cell.createComp(this.sheet, this.block, parm);
			
			//because of overflow logic, we must maintain overflow link from overhead
			//copy over flow attrbute overby and overhead,
			if (tempcell) {
				if (tempcell.overhead) ctrl.overby = tempcell;
				else if(tempcell.overby) ctrl.overby = tempcell.overby;
			}
			this.pushCellI(ctrl, index + i);
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
			size = cells.length;
		jq(this.comp).attr("z.r", newrow);
		for (var i = 0; i < size; i++)
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
			jq(cell.comp).remove();
			cell.cleanup();
		}
			
		this.shiftCellInfo(index, col);
	}
}, {
	/**
	 * Create row DOM node and return zss.Row object
	 * @return zss.Row
	 */
	createComp: function (sheet, block, parm) {
		var zsh = parm.zsh,
			cmp = document.createElement("div"),
			$n = jq(cmp);
		$n.attr({"zs.t": "SRow", "z.r": parm.row}).addClass("zsrow" + (zsh ? " zsh" + zsh :""));
		if (zsh)
			$n.attr("z.zsh", zsh);

		return new zss.Row(sheet, block, cmp);
	},
	/**
	 * Returns a row that copy cells from target
	 * @param zss.Row tmprow a target row to copy it's cells
	 * @return zss.Row
	 */
	copyCells: function (tmprow, rowctrl) {
		var cells = tmprow.cells,
			size = cells.length,
			parm = {row: rowctrl.r};
		if (rowctrl.zsh)
			parm.zsh = rowctrl.zsh;

		for (var i = 0; i < size; i++) {
			parm.col = cells[i].c;
			parm.zsw = (cells[i].zsw) ? cells[i].zsw : null;
			parm.edit = cells[i].edit;
			ctrl = zss.Cell.createComp(rowctrl.sheet, rowctrl.block, parm);
			rowctrl.pushCellE(ctrl);
		}
	}
});