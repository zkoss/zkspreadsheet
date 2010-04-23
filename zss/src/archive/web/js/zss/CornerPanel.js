/* CornerPanel.js

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
 * CornerPanel is used for represent the frozen area
 * 
 * Frozen area contains
 * 1. Top panel if has frozen column
 * 2. Left panel if has frozen row
 */
zss.CornerPanel = zk.$extends(zk.Object, {
	$init: function (sheet) {
		this.$supers('$init', arguments);
		var wgt = sheet._wgt,
			cornerPanel = wgt.$n('co'),
			fzr = sheet.frozenRow,
			fzc = sheet.frozenCol;
		this.id = cornerPanel.id;
		this.sheetid = sheet.sheetid;
		this.comp = cornerPanel;
		this.sheet = sheet;
		cornerPanel.ctrl = this;
		
		if (fzc > -1) {
			//initial top panel
			this.topcomp = jq(cornerPanel).children('DIV:first')[0];
			this.tp = new zss.TopPanel(sheet, this.topcomp, true);
		}
	
		if (fzr > -1) {
			//initial left panel
			var top = this.topcomp;
			this.leftcomp = (top ? jq(top).next('DIV')[0] : jq(cornerPanel).children('DIV:first')[0]);
			this.lp = new zss.LeftPanel(sheet, this.leftcomp, true);	
		}
		
		if (this.tp && this.lp) {
			var blockcomp = jq(this.leftcomp).next('DIV')[0];
			this.block = new zss.CellBlockCtrl(sheet, blockcomp, 0, 0);
			this.block.loadByComp(blockcomp);

			var selareacmp = jq(blockcomp).next("DIV")[0],
				selchgcmp = jq(selareacmp).next("DIV")[0],
				focuscmp = jq(selchgcmp).next("DIV")[0],
				hlcmp = jq(focuscmp).next("DIV")[0];
	
			this.selArea = new zss.SelAreaCtrlCorner(sheet, selareacmp, sheet.initparm.selrange.clone());
			this.selChgArea = new zss.SelChgCtrlCorner(sheet, selchgcmp);
			this.focusMark = new zss.FocusMarkCtrlCorner(sheet, focuscmp, sheet.initparm.focus.clone());
			this.hlArea = new zss.HighlightCorner(sheet, hlcmp, sheet.initparm.hlrange.clone(), "inner");
		}
	},
	cleanup: function () {
		this.invalid = true;
		
		if (this.tp) {
			this.tp.cleanup();
			this.tp = null;
			this.topcomp = null;
		}
		
		if (this.lp) {
			this.lp.cleanup();
			this.lp = null;
			this.leftcomp = null;
		}
		
		if (this.block) {
			this.block.cleanup();
			this.block = null;
		}
		
		if (this.selArea) {
			this.selArea.cleanup();
			this.selArea = null;
		}
		
		if (this.selChgArea) {
			this.selChgArea.cleanup();
			this.selChgArea = null;
		}
		if (this.focusMark) {
			this.focusMark.cleanup();
			this.focusMark = null;
		}
		if (this.hlArea) {
			this.hlArea.cleanup();
			this.hlArea = null;
		}
		
		if(this.comp) this.comp.ctrl = null;
		this.comp = null;
		this.sheet = null;
	},
	_fixSize: function () {
		var sheet = this.sheet,
			wgt = sheet._wgt,
			lw = sheet.leftWidth,
			th = sheet.topHeight;
			leftw = lw - 1,
			toph = th - 1,
			fzr = sheet.frozenRow,
			fzc = sheet.frozenCol;
			
		if (fzr > -1)
			toph = toph + sheet.custRowHeight.getStartPixel(fzr + 1);

		if (fzc > -1)
			leftw = leftw + sheet.custColWidth.getStartPixel(fzc + 1);
		var name = "#" + this.sheetid,
			sid = this.sheetid + "-sheet";
		
		zcss.setRule(name + " .zscorner", ["width", "height"],
			[(fzc> -1 ? leftw : leftw + 1) + "px", (fzr >-1 ? toph : toph + 1) + "px"], true, sid);
	},
	/**
	 * Sets the column's width position index
	 * @param int col the column
	 * @param int zsw the width position index
	 */
	appendZSW: function (col, zsw) {
		if (this.block) this.block.appendZSW(col, zsw);
		if (this.tp) this.tp.appendZSW(col, zsw);
		if (this.lp) this.lp.appendZSW(col, zsw);
	},
	/**
	 * Sets the row's height position index
	 * @param int row the row
	 * @param int the height position index
	 */
	appendZSH: function (row, zsh) {
		if (this.block) this.block.appendZSH(row, zsh);
		if (this.tp) this.tp.appendZSH(row, zsh);
		if (this.lp) this.lp.appendZSH(row, zsh);
	},
	/**
	 * Insert new column
	 * @param int col the column
	 * @param size the size of the column
	 * @param extnm
	 */
	insertNewColumn: function (col, size, extnm) {
		if (this.block) this.block.insertNewColumn(col, size);
		if (this.tp) this.tp.insertNewColumn(col, size, extnm);
		if (this.lp) this.lp.insertNewColumn(col, size);
	},
	/**
	 * Insert new row
	 * @param int row the row
	 * @param int size the size of the row
	 * @param extnm 
	 */
	insertNewRow: function(row, size, extnm) {
		if (this.block) this.block.insertNewRow(row, size);
		if (this.tp) this.tp.insertNewRow(row, size);
		if (this.lp) this.lp.insertNewRow(row, size, extnm);
	},
	/**
	 * Remove column
	 * @param int col the column
	 * @param int size the size the of the column
	 * @param extnm
	 */
	removeColumn: function (col, size, extnm) {
		if (this.block) this.block.removeColumn(col, size);
		if (this.tp) this.tp.removeColumn(col, size, extnm);
		if (this.lp) this.lp.removeColumn(col, size);
	},
	/**
	 * Remove row
	 * @param int row the row
	 * @param int size the size if the row
	 * @param extnm
	 */
	removeRow: function (row, size, extnm) {
		if (this.block) this.block.removeRow(row, size);
		if (this.tp) this.tp.removeRow(row, size);
		if (this.lp) this.lp.removeRow(row, size, extnm);
	}
});