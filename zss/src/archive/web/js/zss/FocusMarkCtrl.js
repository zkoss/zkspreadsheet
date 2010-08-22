/* FocusMarkCtrl.js

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
 * FocusMarkCtrl control the focus mark of the spreadsheet
 */
zss.FocusMarkCtrl = zk.$extends(zk.Object, {
	$init: function (sheet, cmp, pos) {
		this.$supers('$init', arguments);
		this.id = cmp.id;
		this.sheetid = sheet.sheetid;
		this.comp = cmp;
		var local = this;
		
		this.sheet = sheet;
		this.row = pos.row;
		this.column = pos.column;
		
		cmp.ctrl = this;
	},
	cleanup: function () {
		this.invalid = true;

		if (this.comp) this.comp.ctrl = null;
		this.comp = this.sheet =  null;
	},
	/**
	 * Locate the focus mark to cell
	 * @param int row row index
	 * @param int col column index
	 */
	relocate: function (row, col) {
		var sheet = this.sheet;
		this.row = row;
		this.column = col;
		var cell = sheet.getCell(row, col),
			mr = cell && cell.merr ? cell.merr : col,
			custColWidth = sheet.custColWidth,
			custRowHeight = sheet.custRowHeight,
			l = custColWidth.getStartPixel(col),
			t = custRowHeight.getStartPixel(row),
			w = custColWidth.getStartPixel(mr + 1) - l,
			h = custRowHeight.getSize(row);
		
		this.relocate_(l, t, w, h);
	},
	/**
	 * Sets the focus mark position
	 * @param int left position
	 * @param int top position
	 * @param int width
	 * @param int height
	 */
	relocate_: function(l, t, w, h) {
		var sheet = this.sheet,
			dp = sheet.dp;
		
		l += sheet.leftWidth;//adjust to block position.
		t += sheet.topHeight;//adjust to block position.
		jq(this.comp).css({'width': jq.px0(w - 3), 'height': jq.px0(h - 3), 'left': jq.px(l - 2), 'top': jq.px(t - 2)});
	},
	/**
	 * Display focus mark
	 */
	showMark: function () {
		this.comp.style.display = "block";
	},
	/**
	 * Hide focus mark
	 */
	hideMark: function () {
		this.comp.style.display = "none";
	}
});

/**
 * FocusMarkCtrlCorner control the focus mark at corner panel
 */
zss.FocusMarkCtrlCorner = zk.$extends(zss.FocusMarkCtrl, {});

/**
 * FocusMarkCtrlLeft control the focus mark at left panel
 */
zss.FocusMarkCtrlLeft = zk.$extends(zss.FocusMarkCtrl, {
	//override
 	relocate_: function (l, t, w, h) {
		var sheet = this.sheet;
		l += sheet.leftWidth - 1;
		if (sheet.lp.toppad) {
			t -= sheet.lp.toppad;
		}
		jq(this.comp).css({'width': jq.px0(w - 3), 'height': jq.px0(h - 3), 'left': jq.px(l - 2), 'top': jq.px(t - 2)});
	}
});

/**
 *  FocusMarkCtrlTop control the focus mark at the top panel
 */
zss.FocusMarkCtrlTop = zk.$extends(zss.FocusMarkCtrl, {
	//override
  	relocate_: function(l, t, w, h) {
		var sheet = this.sheet;
		t += sheet.topHeight - 1;//adjust to block position.
		jq(this.comp).css({'width': jq.px0(w - 3), 'height': jq.px0(h - 3), 'left': jq.px(l - 2),'top': jq.px(t - 2)});
	}
});