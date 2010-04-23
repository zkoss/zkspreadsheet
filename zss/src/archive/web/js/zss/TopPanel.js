/* TopPanel.js

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
 * TopPanel represent the top rectangle area of the spreadsheet. It contains column headers of the spreadsheet
 */
zss.TopPanel = zk.$extends(zk.Object, {
	$init: function (sheet, node, corner) {
		this.$supers('$init', arguments);
		var wgt = sheet._wgt,
			inner = jq(node).children('DIV:first')[0],
			head = jq(inner).children('DIV:first')[0],
			fontSize = corner && zk.ie ? wgt._getTopHeaderFontSize() : null;//head: contains headers.

		if (fontSize)
			this.fontSize = fontSize;
		this.id = node.id;
		this.sheetid = sheet.sheetid;
		this.comp = node;
		
		zk(node).disableSelection();//disable selectable
		this.icomp = inner;
		this.hcomp = head;
		
		this.hidehead = ((jq(head).attr('z.hide') == "true") ? true : false);
		this.sheet = sheet;
		node.ctrl = this;
		
		
		this.headers = this._initHeader(sheet, head, fontSize);
	
		zkS.trimFirst(inner, "DIV");
		zkS.trimLast(inner, "DIV");
		zkS.trimFirst(head, "DIV");
		zkS.trimLast(head, "DIV");

		if (!corner) {
			this._initFrozenRow(sheet, head);
		}

		wgt.domListen_(node, 'onMouseOver', '_doTopPanelMouseOver');
		wgt.domListen_(node, 'onMouseOut', '_doTopPanelMouseOut');
	},
	_initHeader: function (sheet, head, fontSize) {
		var headers = [],
			nodes = jq(head).children('div'),
			size = nodes.length,
			header,
			idx,
			boundary;
		for (var i = 0; i < size; i++) {
			header = nodes[i];
			if (jq(header).attr('zs.t') == 'STheader') {
				if (fontSize)
					jq(header).css('font-size', fontSize);
				idx = zk.parseInt(jq(header).attr('z.c'));
				boundary = nodes[i + 1];
				headers.push(new zss.Header(sheet, header, boundary, idx, zss.Header.HOR));
			}
		}
		return headers;
	},
	_initFrozenRow: function (sheet, head) {
		var fzr = sheet.frozenRow;
		if (fzr > -1) {
			var topBlock = jq(head).next('DIV')[0];//topBlock: contains freeze Rows
			this.block = new zss.CellBlockCtrl(sheet, topBlock, 0, 0);

			this.block.loadByComp(topBlock);
			var selArea = jq(topBlock).next('DIV')[0],
				selChg = jq(selArea).next('DIV')[0],
				focus = jq(selChg).next('DIV')[0],
				highlight = jq(focus).next('DIV')[0];
			this.selArea = new zss.SelAreaCtrlTop(sheet, selArea, sheet.initparm.selrange.clone());
			this.selChgArea = new zss.SelChgCtrlTop(sheet, selChg);
			this.focusMark = new zss.FocusMarkCtrlTop(sheet, focus, sheet.initparm.focus.clone());
			this.hlArea = new zss.HighlightTop(sheet, highlight, sheet.initparm.hlrange.clone(), "inner");
		}
	},
	/**
	 *  IE need font size to display correctly in corner panel
	 *  @return string font size of the header , or null means this's not in corner panel
	 */
	_getCornerHeaderFontSize: function () {
		return this.fontSize;
	},
	cleanup: function () {
		var wgt = this.sheet._wgt,
			n = this.hcomp;
		wgt.domUnlisten_(n, 'onMouseOver', '_doTopPanelMouseOver');
		wgt.domUnlisten_(n, 'onMouseOut', '_doTopPanelMouseOut');
		
		this.invalid = true;
		if (this.comp) this.comp.ctrl = null;
		this.comp = this.icomp = this.hcomp = this.sheet = null;
		
		if (this.fontSize)
			this.fontSize = null;
		
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
	_clearAllHeader: function () {
		jq(this.hcomp).text('');
		var i = this.headers.length;
		while (i--) {
			this.headers[i].cleanup();
			this.headers[i] = null;
		}
		this.headers = [];
	},
	_doMouseover: function (evt) {
		if (this.sheet.headerdrag) return;
		var n = evt.domTarget;
		if (jq(n).attr('zs.t') == "SBoun")
			n.parentNode.ctrlref._processDrag(true);
	},
	_doMouseout: function (evt){
		if (this.sheet.headerdrag) return;
		var n = evt.domTarget;
		if (jq(n).attr('zs.t') == "SBoun")
			n.parentNode.ctrlref._processDrag(false);
	},
	_createEastHeader: function (headerdata, width) {
		if (this.hidehead) return;
		
		for (var j = 0; j < width; j++) {
			var header = headerdata[j],
				parm = {type: zss.Header.HOR};
			zkS.copyParm(header, parm, ["ix", "nm", "zsw", "zsh"]);
			var headerctrl = zss.Header.createComp(this.sheet, parm);
			this.pushHeaderE(headerctrl);
		}

		this._updateSelectionCSS();
	},
	_createWestHeader: function (headerdata, width) {
		if (this.hidehead) return;
		for (var j = 0; j < width; j++) {
			var header = headerdata[j],
				parm = {type: zss.Header.HOR};
			zkS.copyParm(header, parm, ["ix", "nm", "zsw", "zsh"]);
			var headerctrl = zss.Header.createComp(this.sheet, parm);
			this.pushHeaderS(headerctrl);
		}
		this._updateSelectionCSS();
	},
	_createJumpHeader: function (headerdata, width) {
		if (this.hidehead) return;
		//clear all
		this._clearAllHeader();
		
		var start = new Date().getTime();
		for (var j = 0; j < width; j++) {
			var header = headerdata[j],
				parm = {type: zss.Header.HOR};
			zkS.copyParm(header, parm, ["ix", "nm", "zsw", "zsh"]);
			var headerctrl = zss.Header.createComp(this.sheet, parm);
			this.pushHeaderE(headerctrl);
		}

		this._updateSelectionCSS();
	},
	/**
	 * Sets the header to the end
	 * @param zss.Header headerctrl
	 */
	pushHeaderE: function (headerctrl) {
		if (this.hidehead) return;
		this.headers.push(headerctrl);
		this.hcomp.appendChild(headerctrl.comp);
		this.hcomp.appendChild(headerctrl.bcomp);
	},
	/**
	 * Sets the header to start
	 * @param zss.Header headerctrl
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
			throw('index out of bound:'+ index + ' > ' + size);

		if (index == 0)
			this.pushHeaderS(headerctrl);
		else if (index == size)
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
	_removeWestHeader: function (size) {
		if(this.hidehead) return;
		var headers = this.headers,
			header,
			hcomp = this.hcomp;
		while (size--) {
			header = headers.shift();
			jq(header.comp).remove();
			jq(header.bcomp).remove();
			header.cleanup();
		}
	},
	_removeEastHeader: function (size){
		if (this.hidehead) return;
		var headers = this.headers,
			header,
			hcomp = this.hcomp;
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
	},
	_updateWidth: function (width) {
		jq(this.comp).css('width', jq.px(width));
		this.width = width;	
		this._updateBlockWidth();
	},
	_updateLeftPos: function (pos) {
		jq(this.icomp).css('left', jq.px(pos));
		this.leftpos = pos;
		this._updateBlockWidth();
	},
	_updateLeftPadding: function (leftpad) {
		leftpad = leftpad - this.sheet.leftWidth;
		jq(this.icomp).css('padding-left', jq.px(leftpad));
		this.leftpad = leftpad;
		this._updateBlockWidth();
	},
	_updateBlockWidth: function () {
		if (!this.block) return;
		var width = this.width,
			leftpos = this.leftpos,
			leftpad = this.leftpad;
		width = width - (leftpos ? leftpos : 0) - (leftpad ? leftpad : 0);
		if (width < 0) width = 0;
		jq(this.block.comp).css('width', jq.px(width));
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
				jq(header.comp)[remove ? 'removeClass' : 'addClass']("zstop-sel");
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

		var left = this.headers[0].index,
			right = left + this.headers.length - 1;
		if (left > col || right < col) return;
		
		var index = col - left,
			header = this.headers[index];
		header.appendZSW(zsw);
	},
	/**
	 * Sets the height position index
	 * @param row row index
	 * @param int the height position index
	 */
	appendZSH: function (row, zsh) {
		if (this.block)
			this.block.appendZSH(row, zsh);
	},
	/**
	 * Insert new column
	 * @param int col column
	 * @param int size the size of the column
	 * @param array extnm
	 */
	insertNewColumn: function (col, size, extnm) {
		if (this.block)
			this.block.insertNewColumn(col,size);

		var left = this.headers[0].index,
			right = left + this.headers.length -1; 
		if (col > (right+1) || col < left) return;
		
		var index = col - left,
			ctrl,
			nm,
			parm = {type: zss.Header.HOR},
			fontSize = zk.ie ? this._getCornerHeaderFontSize() : null;
		for (var i = 0; i < size; i++) {
			parm.ix = col + i;
			parm.nm = extnm[i]
			ctrl = zss.Header.createComp(this.sheet, parm);

			if (fontSize)
				jq(ctrl.getHeaderNode()).css('font-size', fontSize);

			this.pushHeaderI(ctrl, index + i);
		}
		//var tail = headers.slice(index,size);
		extnm = extnm.slice(size, extnm.length);
		this.shiftHeaderInfo(index + size, col + size,extnm);
	},
	/**
	 * Insert new row
	 * @param int row row index
	 * @param int size the size of the row
	 */
	insertNewRow: function (row, size) {
		if (this.block)
			this.block.insertNewRow(row,size);
	},
	/**
	 * Shift header
	 * @param int index start shift index
	 * @param int newcol new column index
	 * @param array extnm
	 */
	shiftHeaderInfo: function (index, newcol, extnm) {
		var size = this.headers.length,
			j = 0;
		for (var i = index; i < size; i++) {
			if (!extnm[j]) zk.log("undefined header to assing>>"+(newrow+j),"always");
			this.headers[i].resetInfo(newcol + j, extnm[j]);
			j++;
		}
	},
	/**
	 * Remove column
	 * @param int col column index
	 * @param int size the size of the column
	 * @param array extnm
	 */
	removeColumn: function (col, size, extnm) {
		if (this.block)
			this.block.removeColumn(col,size);

		var left = this.headers[0].index,
			right = left + this.headers.length -1; 
		if (col > (right+1) || col < left) return;
		
		var index = col - left;
		if ((col + size) > right)
			size = right - col + 1;
		this._removeHeader(index, size);
		this.shiftHeaderInfo(index, col, extnm);
	},
	/**
	 * Remove row
	 * @param int row row index
	 * @param int size
	 * @param array extnm
	 */
	removeRow: function (row, size) {
		if (this.block)
			this.block.removeRow(row, size);
	},
	_updateSelectionCSS: function () {
		var sheet = this.sheet,
			selRange = sheet.selArea.lastRange;
		if (selRange) {
			var left = selRange.left,
				right = selRange.right;
			this.updateSelectionCSS(left, right, false || sheet.state == zss.SSheetCtrl.NOFOCUS);
		}
	},
	_fixSize: function() {
		var sheet = this.sheet,
			th = sheet.topHeight,
			toph = th - 1,
			fzr = sheet.frozenRow;
		
		if (fzr > -1)
			toph = toph + sheet.custRowHeight.getStartPixel(fzr + 1);
		
		var name = "#" + this.sheetid,
			sid = this.sheetid + "-sheet";
		zcss.setRule(name + " .zstop", ["height"], [(fzr > -1 ? toph - 1: toph) + "px"], true, sid);
		zcss.setRule(name + " .zstopi", ["height"], [toph + "px"], true, sid);
	}
}, {});