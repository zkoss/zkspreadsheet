/* LeftPanel.js

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
 * LeftPanel represent the left area of the spreadsheet. It contains row headers of the spreadsheet
 */
zss.LeftPanel = zk.$extends(zk.Object, {
	$init: function (sheet, node, corner) {
		this.$supers('$init', arguments);	
		var wgt = sheet._wgt,
			pad = jq(node).children('DIV:first')[0],
			inner = jq(pad).next('DIV')[0],
			head = jq(inner).children('DIV:first')[0];

		this.id = node.id;
		this.sheetid = sheet.sheetid;
		this.comp = node;
		zk(node).disableSelection();//disable selectable
		
		this.padcomp = pad;
		this.icomp = inner;
		this.hcomp = head;
		this.hidehead = ((jq(head).attr("z.hide") == "true") ? true : false);
		
		this.sheet = sheet;
		node.ctrl = this;
		
		this.headers = this._initHeader(sheet, head);
		
		//not a corner left panel
		if (!corner) {
			this._initFrozenColumn(sheet, head);
		}

		wgt.domListen_(node, 'onMouseOver', '_doLeftPanelMouseOver');
		wgt.domListen_(node, 'onMouseOut', '_doLeftPanelMouseOut');
	},
	_initHeader: function (sheet, head) {
		var headers = [],
			nodes = jq(head).children('div'),
			size = nodes.length,
			header,
			idx,
			boundary;

		for (var i = 0; i < size; i++) {
			header = nodes[i];
			if (jq(header).attr('zs.t') == 'SLheader') {
				idx = zk.parseInt(jq(header).attr('z.r'));
				boundary = nodes[i + 1];
				headers.push(new zss.Header(sheet, header, boundary, idx, zss.Header.VER));
			}
		}
		return headers;
	},
	_initFrozenColumn: function (sheet, head) {
		var fzc = sheet.frozenCol;
		if (fzc > -1) {
			var leftBlock = jq(head).next("DIV")[0]; //leftBlock: contains freeze Columns
			this.block = new zss.CellBlockCtrl(sheet, leftBlock, 0, 0);
			this.block.loadByComp(leftBlock);
			
			var selArea = jq(leftBlock).next("DIV")[0],
				selChg = jq(selArea).next("DIV")[0],
				focus = jq(selChg).next("DIV")[0],
				highlight = jq(focus).next("DIV")[0];
			
			this.selArea = new zss.SelAreaCtrlLeft(sheet, selArea, sheet.initparm.selrange.clone());
			this.selChgArea = new zss.SelChgCtrlLeft(sheet, selChg);
			this.focusMark = new zss.FocusMarkCtrlLeft(sheet, focus, sheet.initparm.focus.clone());
			this.hlArea = new zss.HighlightLeft(sheet, highlight, sheet.initparm.hlrange.clone(), "inner");

		}
	},
	cleanup: function () {
		var wgt = this.sheet._wgt,
			n = this.comp;
		wgt.domUnlisten_(n, 'onMouseOver', '_doLeftPanelMouseOver');
		wgt.domUnlisten_(n, 'onMouseOut', '_doLeftPanelMouseOut');
		
		this.invalid = true;
		if (this.comp) this.comp.ctrl = null;
		this.comp = this.icomp = this.hcomp = this.sheet = null;
		
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
		
		var i = this.headers.length;
		while (i--) {
			this.headers[i].cleanup();
			this.headers[i] = null;
		}
		this.headers = null;
	},
	_doMouseover: function (evt) {
		var n = evt.domTarget;
		if (jq(n).attr('zs.t') == "SBoun")
			n.parentNode.ctrlref._processDrag(true, false);
		if (jq(n).attr('zs.t') == "SBoun2")
			n.parentNode.ctrlref._processDrag(true, true);
	},
	_doMouseout: function (evt) {
		var n = evt.domTarget;
		if (jq(n).attr('zs.t') == "SBoun")
			n.parentNode.ctrlref._processDrag(false, false);
		if (jq(n).attr('zs.t') == "SBoun2")
			n.parentNode.ctrlref._processDrag(false, true);
	},
	_clearAllHeader: function () {
		jq(this.hcomp).text('');
		var i = this.headers.length
		while (i--) {
			this.headers[i].cleanup();
			this.headers[i] = null;
		}
		this.headers = [];
	},
	_createSouthHeader: function (headerdata, height) {
		if (this.hidehead) return;
		for (var j = 0;j < height; j++) {
			var header = headerdata[j],
				parm = {type: zss.Header.VER};
			zkS.copyParm(header, parm, ["ix", "nm", "zsw", "zsh"]);
			var headerctrl = zss.Header.createComp(this.sheet, parm);
			this.pushHeaderE(headerctrl);
		}
		this._updateSelectionCSS();
	},
	_createNorthHeader: function (headerdata, height) {
		if (this.hidehead) return;
		for (var j = 0; j < height; j++) {
			var header = headerdata[j],
				parm = {type: zss.Header.VER};
			zkS.copyParm(header, parm, ["ix", "nm", "zsw", "zsh"]);
			var headerctrl = zss.Header.createComp(this.sheet, parm);
			this.pushHeaderS(headerctrl);
		}
		this._updateSelectionCSS();
	},
	_createJumpHeader: function (headerdata, height) {
		if (this.hidehead) return;
		//clear all
		this._clearAllHeader();
		
		for (var j = 0; j < height; j++) {
			var header = headerdata[j],
				parm = {type: zss.Header.VER};
			zkS.copyParm(header, parm, ["ix", "nm", "zsw", "zsh"]);
			var headerctrl = zss.Header.createComp(this.sheet, parm);
			this.pushHeaderE(headerctrl);
		}
		this._updateSelectionCSS();
	},
	/**
	 * Sets the header to the end
	 * @param zss.Header
	 */
	pushHeaderE: function (headerctrl) {
		if (this.hidehead) return;
		this.headers.push(headerctrl);
		this.hcomp.appendChild(headerctrl.comp);
		this.hcomp.appendChild(headerctrl.bcomp);
	},
	/**
	 * Sets the header to start
	 * @param zss.Header
	 */
	pushHeaderS: function (headerctrl) {
		if (this.hidehead) return;
		this.headers.unshift(headerctrl);
		this.hcomp.insertBefore(headerctrl.bcomp, this.hcomp.firstChild);
		this.hcomp.insertBefore(headerctrl.comp, headerctrl.bcomp);
	},
	/**
	 * Sets the header position by index
	 * @param zss.Header
	 * @param int index
	 */
	pushHeaderI: function (headerctrl, index) {
		if (this.hidehead) return;
		var headers = this.headers,
			size = headers.length;
		if (index > size)
			throw('index out of bound:' + index + ' > ' + size);

		if (index == 0)
			this.pushHeaderS(headerctrl);
		else if(index == size)
			this.pushHeaderE(headerctrl);
		else {
			var tail = headers.slice(index, size);
			headers.length = index;
			headers.push(headerctrl);
			headers.push.apply(headers, tail);
			
			this.hcomp.insertBefore(headerctrl.bcomp, tail[0].comp);
			this.hcomp.insertBefore(headerctrl.comp, headerctrl.bcomp);
		}
	},
	_removeNorthHeader: function (size) {
		if (this.hidehead) return;

		var headers = this.headers,
			header;
		while (size--) {
			header = headers.shift();
			jq(header.comp).remove();
			jq(header.bcomp).remove();
			header.cleanup();
		}
	},
	_removeSouthHeader: function (size) {
		if (this.hidehead) return;
		var headers = this.headers,
			header;
		while (size--) {
			header = headers.pop();
			jq(header.comp).remove();
			jq(header.bcomp).remove();
			header.cleanup();
		}
	},
	_removeHeader: function (index, size) {
		var ctrl,
			headers = this.headers;
		
		if (index > headers.length) return;
		
		if (index == 0)
			this._removeNorthHeader(size);
		else {
			var rem = headers.slice(index, index + size),
				tail = headers.slice(index + size, headers.length);
			headers.length = index;
			headers.push.apply(headers, tail);
			
			var header = rem.pop();
			for (; header; header = rem.pop()) {
				jq(header.comp).remove();
				jq(header.bcomp).remove();
				header.cleanup();
			}
		}
	},
	_updateHeight: function (height) {
		if (this.height == height) return;
		jq(this.comp).css('height', jq.px0(height));
		this.height = height;
		this._updateBlockHeight();
	},	
	_updateTopPos: function (toppos) {
		if (this.toppos == toppos) return;
		jq(this.icomp).css('top', jq.px(toppos));
		this.toppos = toppos;
		this._updateBlockHeight();
	},
	_updateTopPadding: function (toppad) {
		if (this.toppad == toppad) return;
		jq(this.padcomp).css('height', jq.px0(toppad));
		this.toppad = toppad;	
		
		if (this.selArea)
			this.selArea.relocate();

		if (this.selChgArea)
			this.selChgArea.relocate();

		if (this.focusMark) {
			var pos = this.sheet.getLastFocus();
			this.focusMark.relocate(pos.row, pos.column);
		}
		this._updateBlockHeight();
	},
	_updateBlockHeight: function () {
		if (!this.block) return;
		var height = this.height,
			toppos = this.toppos,
			toppad = this.toppad;
		height = height - (toppos ? toppos : 0) - (toppad ? toppad : 0);
		if (height < 0) height = 0;
		jq(this.block.comp).css('height', jq.px0(height));
	},
	/**
	 * Sets the selection range of the header
	 * @param int from row selection start index
	 * @param int to row selection end index
	 */
	updateSelectionCSS: function (from, to, remove) {
		var i = this.headers.length,
			header;
		while (i--) {
			header = this.headers[i];
			if (header.index >= from && header.index <= to)
				jq(header.comp)[remove ? 'removeClass' : 'addClass']("zsleft-sel");
		}
	},
	_updateSelectionCSS: function () {
		var sheet = this.sheet,
			selRange = sheet.selArea.lastRange;
		if (selRange) {
			var top = selRange.top,
				bottom = selRange.bottom;
			this.updateSelectionCSS(top, bottom, false || sheet.state == zss.SSheetCtrl.NOFOCUS);
		}
	},
	/**
	 * Sets the width position index
	 * @param int col column index
	 * @param int zsw the width position index
	 */
	appendZSW: function (col, zsw) {
		if (this.block)
			this.block.appendZSW(col, zsw);
	},
	/**
	 * Sets the height position index
	 * @param row row index
	 * @param int the height position index
	 */
	appendZSH: function (row, zsh) {
		if (this.block)
			this.block.appendZSH(row, zsh);
		var top = this.headers[0].index,
			bottom = top + this.headers.length - 1; 
		if (top > row || bottom < row) return;
		
		var index = row - top,
			header = this.headers[index];
		header.appendZSH(zsh);
	},
	/**
	 * Insert new column
	 * @param int col column
	 * @param int size the size of the column
	 */
	insertNewColumn: function (col, size) {
		if (this.block)
			this.block.insertNewColumn(col, size);
	},
	/**
	 * Insert new row
	 * @param int row row index
	 * @param int size the size of the row
	 * @param array extnm
	 */
	insertNewRow: function (row, size, extnm) {
		if (this.block)
			this.block.insertNewRow(row, size);
		
		var top = this.headers[0].index,
			bottom = top + this.headers.length - 1; 
		if (row > (bottom + 1) || row < top) return;
		
		var index = row - top,
			ctrl,
			nm,
			parm = {type: zss.Header.VER};
		
		for (var i = 0; i < size; i++) {
			parm.ix = row + i;
			parm.nm = extnm[i];
			ctrl = zss.Header.createComp(this.sheet, parm);
			this.pushHeaderI(ctrl, index + i);
		}
		extnm = extnm.slice(size, extnm.length);
		this.shiftHeaderInfo(index + size, row + size, extnm);
	},
	/**
	 * Shift header
	 * @param int index start shift index
	 * @param array extnm
	 */
	shiftHeaderInfo: function (index, newrow, extnm) {
		var size = this.headers.length,
			j = 0;
		for (var i = index; i < size; i++) {
			if(!extnm[j]) zk.log("undefined header to assing>>"+(newrow+j),"always");
			this.headers[i].resetInfo(newrow + j, extnm[j]);
			j++;
		}
	},
	/**
	 * Remove column
	 * @param int col column index
	 * @param int size the size of the column
	 */
	removeColumn: function (col, size) {
		if (this.block)
			this.block.removeColumn(col, size);
	},
	/**
	 * Remove row
	 * @param int row row index
	 * @param int size
	 * @param array extnm
	 */
	removeRow: function (row, size, extnm) {
		if (this.block)
			this.block.removeRow(row, size);

		var top = this.headers[0].index,
			bottom = top + this.headers.length - 1; 
		if(row > (bottom+1) || row < top) return;
		
		var index = row - top;
		if ((row + size) > bottom)
			size = bottom - row + 1;

		this._removeHeader(index, size);
		this.shiftHeaderInfo(index, row, extnm);
	},
	_fixSize: function() {
		var sheet = this.sheet,
			lw = sheet.leftWidth,
			leftw = lw-1,
			fzc = sheet.frozenCol;
		
		if (fzc > -1)
			leftw = leftw + sheet.custColWidth.getStartPixel(fzc + 1);
		
		var name = "#" + this.sheetid,
			sid = this.sheetid + "-sheet";

		zcss.setRule(name + " .zsleft", ["width"], [(fzc > -1 ? leftw - 1 : leftw) + "px"], true, sid);
		zcss.setRule(name + " .zslefti", ["width"], [leftw + "px"], true, sid);
	}
});