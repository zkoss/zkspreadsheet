/* AreaCtrl.js

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

(function () {
	/**
	 * Clone the opacity dom element and replace.
	 * This is for IE 8. when opacity area large than a limit, the mask of opacity will fail.
	 * http://yuilibrary.com/projects/yui2/ticket/2528621
	 */
	function _redrawOpacity (obj) {
		var inner = obj.icomp,
			cloneInner = jq(inner).clone()[0];
		jq(inner).replaceWith(cloneInner);
		obj.icomp = cloneInner;
	}

	function _showSelectionArea (obj, redrawOpacity) {
		var range = obj.lastRange;
		if (!range) return;
		jq(obj.comp).css('display', 'block');
		jq(obj.icomp)[((range.top != range.bottom) || (range.left != range.right)) ? 'addClass' : 'removeClass']("zsselecti-r");
		if (zk.ie8 && redrawOpacity)
			_redrawOpacity(obj);
	}
/**
 * Locate selection area.
 */
zss.AreaCtrl = zk.$extends(zk.Object, {
	$init: function (sheet, cmp, range, mode) {
		this.$supers('$init', arguments);
		this.id = cmp.id;
		this.sheetid = sheet.sheetid;
		cmp.ctrl = this;
		this.comp = cmp;
		this.icomp = jq(cmp).children("DIV:first")[0];
		this.sheet = sheet;
		this.lastRange = range ? range : new zss.Range(0,0,0,0);
		this.mode = mode;//null,inner,outer
	},
	cleanup: function () {
		this.invalid = true;
		if (this.comp) this.comp.ctrl = null;
		this.comp = this.icomp = this.sheet = null;
	},
	/**
	 * Locate selection area
	 * @param zss.Range range of the selection
	 */
	relocate: function (range) {
		var sheet = this.sheet;
		if (!range) {
			range = this.lastRange;
			if (!range) return;
		} else
			this.lastRange = range;

		var custColWidth = sheet.custColWidth,
			custRowHeight = sheet.custRowHeight,
			l = custColWidth.getStartPixel(range.left),
			t = custRowHeight.getStartPixel(range.top),
			w = custColWidth.getStartPixel(range.right + 1) - l,
			h = custRowHeight.getStartPixel(range.bottom + 1) - t;
		
		this.relocate_(l, t, w, h);
	},
	relocate_: function(l, t, w, h) {

		var sheet = this.sheet,
			dp = sheet.dp;
		
		l += sheet.leftWidth;//adjust to block position.
		t += sheet.topHeight;//adjust to block position.
		l = l - 2;
		t = t - 2;
		w = w - 3;
		h = h - 3;
		if (this.mode == "inner") {
			l = l + 1;
			t = t + 1;
			w = w - 2;
			h = h - 2;
		}

		jq(this.comp).css({'width': jq.px(w), 'height': jq.px(h), 'left': jq.px(l), 'top': jq.px(t)});
	},
	/**
	 * Display selection area
	 */
	showArea: function () {
		var range = this.lastRange;
		if (!range) return;

		jq(this.comp).css('display', 'block');
	},
	/**
	 * Hide selection area 
	 */
	hideArea: function () {
		jq(this.comp).css('display', 'none');
	}
});

/**
 * Locate data panel selection area
 */
zss.SelAreaCtrl = zk.$extends(zss.AreaCtrl, {
	//override
	showArea: function () {
		_showSelectionArea(this, true);
	}
});

/**
 * Locate data panel selection area
 */
zss.SelChgCtrl = zk.$extends(zss.AreaCtrl, {});

/**
 * Locate corner panel selection area
 */
zss.AreaCtrlCorner = zk.$extends(zss.AreaCtrl, {});

/**
 * Locate corner panel selection area
 */
zss.SelAreaCtrlCorner = zk.$extends(zss.AreaCtrlCorner, {
	//override
	showArea: function () {
		_showSelectionArea(this);
	}
});

/**
 *  Locate corner panel selection change area
 */
zss.SelChgCtrlCorner = zk.$extends(zss.AreaCtrlCorner, {});

/**
 * Locate left panel selection area
 */
zss.AreaCtrlLeft = zk.$extends(zss.AreaCtrl, {
	//override
  	relocate_: function (l, t, w, h) {
		var sheet = this.sheet;
		l += sheet.leftWidth - 1;
		
		if (sheet.lp.toppad) {
			t -= sheet.lp.toppad;
		}
		
		l = l - 2;
		t = t - 2;
		w = w - 3;
		h = h - 3;
		if (this.mode == "inner") {
			l = l + 1;
			t = t + 1;
			w = w - 2;
			h = h - 2;
		}

		jq(this.comp).css({'width': jq.px(w), 'height': jq.px(h), 'left': jq.px(l),'top' : jq.px(t)});
	}
});

/**
 * Locate left panel header selection area
 */
zss.SelAreaCtrlLeft = zk.$extends(zss.AreaCtrlLeft, {
	//override
	showArea: function () {
		_showSelectionArea(this);
	}
});

/**
 * Locate left panel header selection change area
 */
zss.SelChgCtrlLeft = zk.$extends(zss.AreaCtrlLeft, {});

/**
 * Locate top panel selection area
 */
zss.AreaCtrlTop = zk.$extends(zss.AreaCtrl, {
	//override
  	relocate_: function(l , t, w, h) {
		var sheet = this.sheet;
		t += sheet.topHeight - 1;//adjust to block position.

		l = l - 2;
		t = t - 2;
		w = w - 3;
		h = h - 3;
		if (this.mode == "inner") {
			l = l + 1;
			t = t + 1;
			w = w - 2;
			h = h - 2;
		}

		jq(this.comp).css({'width': jq.px(w), 'height': jq.px(h), 'left': jq.px(l), 'top': jq.px(t)});
	}
});

/**
 * Locate top panel header selection area
 */
zss.SelAreaCtrlTop = zk.$extends(zss.AreaCtrlTop, {
	//override
	showArea: function () {
		_showSelectionArea(this);
	}
});

/**
 *  Locate top panel header selection change area
 */
zss.SelChgCtrlTop = zk.$extends(zss.AreaCtrlTop, {});
})();