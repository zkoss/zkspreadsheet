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
zss.CornerPanel = zk.$extends(zk.Widget, {
	widgetName: 'CornerPanel',
	$o: zk.$void, //owner, fellows relationship no needed
	$init: function (sheet, rowHeadHidden, columnHeadHidden, lCol, tRow, rCol, bRow, data) {
		this.$supers(zss.CornerPanel, '$init', []);
		
		this.sheet = sheet;
		
		var r = sheet.frozenRow,
			c = sheet.frozenCol,
			tp, lp;
		if (c > -1) {
			this.appendChild(tp = this.tp = new zss.TopPanel(sheet, columnHeadHidden, 0, c, data, true), true);
		}
		if (r > -1) {
			this.appendChild(lp = this.lp = new zss.LeftPanel(sheet, rowHeadHidden, 0, r, data, true), true);
		}
		if (tp && lp) {
			this.appendChild(this.block = new zss.CellBlockCtrl(sheet, tRow, lCol, bRow, rCol, data, 'corner'), true);
		}
	},
	redraw: function (out) {
		var uid = this.uuid,
			sheet = this.sheet,
			fRow = sheet.frozenRow,
			fCol = sheet.frozenCol;
		out.push('<div id="', uid, '-co" class="zscorner zsfzcorner" zs.t="SCorner">');
		if (this.tp) {
			this.tp.redraw(out);
		}
		if (this.lp) {
			this.lp.redraw(out);
		}
		if (this.block) {
			this.block.redraw(out);
			out.push('<div id="', uid, '-select" class="zsselect" zs.t="SSelect">',
					'<div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>',
					'<div id="', uid, '-selchg" class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>',
					'<div id="', uid, '-focmark" class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>',
					'<div id="', uid, '-highlight" class="zshighlight" zs.t="SHighlight"><div class="zshighlighti"></div></div>');
		}
		out.push('<div class="zscorneri" ></div></div>');
	},
	bind_: function () {
		this.$supers(zss.CornerPanel, 'bind_', arguments);
		
		this.comp = this.$n('co');
		if (this.block) {
			var sheet = this.sheet,
				selareacmp = this.$n('select'),
				selchgcmp = selareacmp.nextSibling,
				focuscmp = selchgcmp.nextSibling,
				hlcmp = focuscmp.nextSibling;

			this.selArea = new zss.SelAreaCtrlCorner(sheet, selareacmp, sheet.initparm.selrange.clone());
			this.selChgArea = new zss.SelChgCtrlCorner(sheet, selchgcmp);
			this.focusMark = new zss.FocusMarkCtrlCorner(sheet, focuscmp, sheet.initparm.focus.clone());
			this.hlArea = new zss.HighlightCorner(sheet, hlcmp, sheet.initparm.hlrange.clone(), "inner");
		}
	},
	unbind_: function () {
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
		this.sheet = this.tp = this.lp = this.block = this.comp = null;
		this.$supers(zss.CornerPanel, 'unbind_', arguments);
	},
	_cornerHeight: function () {
		var sheet = this.sheet,
			th = sheet.topHeight,
			fzr = sheet.frozenRow;
		return (fzr > -1) ?
			(th - 1 + sheet.custRowHeight.getStartPixel(fzr + 1)) :	th;
	},
	_cornerWidth: function () {
		var sheet = this.sheet,
			lw = sheet.leftWidth,
			fzc = sheet.frozenCol;
		return (fzc > -1) ?
			(lw - 1 + sheet.custColWidth.getStartPixel(fzc + 1)) : lw;
	},
	_fixSize: function () {
		var sheetid = this.sheet.sheetid,
			name = "#" + sheetid,
			sid = sheetid + "-sheet";
		
		zcss.setRule(name + " .zscorner", ["width", "height"],
			[this._cornerWidth() + "px", this._cornerHeight() + "px"], true, sid);
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