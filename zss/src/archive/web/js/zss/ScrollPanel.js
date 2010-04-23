/* ScrollPanel.js

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
 * ScrollPanel is used to handle spreadsheet scroll moving event,  
 */
zss.ScrollPanel = zk.$extends(zk.Object, {
	$init: function (sheet) {
		this.$supers('$init', arguments);
		var local = this,
			wgt = sheet._wgt,
			scrollPanel = wgt.$n('sp');
		
		this.id = scrollPanel.id;
		this.sheetid = sheet.sheetid;
		this.sheet = sheet;
		this.comp = scrollPanel;
		this.currentTop = this.currentLeft = this.timerCount = 0;
		this.leftWidth = 28;//left head width
		this.topHeight = 20;//top head height
		this.timerRunning = false;//is timer for scroll running.
		this.lastMove = "";//the last move of scrolling,
		scrollPanel.ctrl = this;

		
		sheet.insertSSInitLater(function() {//datapanel doesn't ready when cell initialing, so invoke later.
			var dtcmp = local.sheet.dp.comp,//zkSDatapanelCtrl._currcmp(local);
				sccmp = local.comp;
			local.minLeft = local._getMaxScrollLeft(dtcmp, sccmp);
			local.minTop = local._getMaxScrollTop(dtcmp, sccmp);
			local.minHeight = dtcmp.offsetHeight;
			local.minWidth = dtcmp.offsetWidth;

			wgt.domListen_(scrollPanel, 'onScroll', '_doScrollPanelScrolling');
			wgt.domListen_(scrollPanel, 'onMouseDown', '_doScrollPanelMouseDown');
		}, false);
	},
	cleanup: function () {
		var wgt = this.sheet._wgt,
			n = wgt.$n('sp');
		wgt.domUnlisten_(n, 'onScroll', '_doScrollPanelScrolling');
		wgt.domUnlisten_(n, 'onMouseDown', '_doScrollPanelMouseDown');
		
		this.invalid = true;
		if(this.comp) this.comp.ctrl = null;
		this.comp = this.sheet = null;
	},
	_doMousedown: function (evt) {
		var cmp = this.comp;
		if(evt.domTarget != cmp) return;
		var data = zkS._getMouseData(evt, cmp);
		
		//calculate is xy in scroll bar
		var clickInHor = (data[1] < zk(cmp).offsetHeight() && data[1] > cmp.clientHeight);
			clickInVer = (data[0] < zk(cmp).offsetWidth() && data[0] > cmp.clientWidth);
		
		if( (clickInHor&&clickInVer) || (!clickInHor&&!clickInVer)) return;
		var sinfo = this.sheet.sinfo;
		sinfo.pinXY(data[0], data[1], clickInHor);
	},
	_doScrolling: function (evt) {
		var sheet = this.sheet,
			dtcmp = sheet.dp.comp,
			sccmp = this.comp,
			scleft = sccmp.scrollLeft,
			sctop = sccmp.scrollTop,
			moveHorizontal = (scleft != this.currentLeft),
			moveVertical = (sctop != this.currentTop),
			moveLeft = scleft < this.currentLeft,
			moveTop = sctop < this.currentTop;
		
		this.currentLeft = scleft;
		this.currentTop = sctop;
		sheet.tp._updateLeftPos(-this.currentLeft);
		sheet.lp._updateTopPos(-this.currentTop);
		if (this._visFlg) {
			//if fire by scroll to visible, don't show info 
			this._visFlg = false;
		} else {
			/*show scroll info*/
			var sinfo = sheet.sinfo;
			if (moveHorizontal) {
				if (!sinfo.visible || !sinfo.horizontal)
					sinfo.pinLocation(true);
				sinfo.showInfoOnDir(true);
			}
			
			if (moveVertical) {
				if (!sinfo.visible || sinfo.horizontal)
					sinfo.pinLocation(false);
				sinfo.showInfoOnDir(false);
			}
			if (moveVertical || moveHorizontal)
				this._trackScrolling(moveVertical);
		}
	},
	_trackScrolling: function (vertical) {
		if (vertical && this.lastMove == "H")
			this._doScroll(false);//fire scroll horizontal
		else if (!vertical && this.lastMove == "V")
			this._doScroll(true);//fire scroll vertical

		this.lastMove = vertical ? "V" : "H";
		var local = this;
		local.timerCount++;
		setTimeout(function() {
			local.timerCount--;
			if (local.timerCount <= 0) {
				local._doScroll(local.lastMove == "V" ? true : false);
				local.lastMove = "";
				local.timerCount = 0;
			}
		}, zkS.scrollTrackingTimeout);
		
	},
	_doScroll: function (vertical) {
		this.sheet.activeBlock.doScroll(vertical);
	},
	_getMaxScrollLeft: function (dtcmp, sccmp) {
		return (dtcmp.offsetWidth - sccmp.offsetWidth) + zss.Spreadsheet.scrollWidth;
	},
	_getMaxScrollTop: function (dtcmp, sccmp) {
		return (dtcmp.offsetHeight - sccmp.offsetHeight) + zss.Spreadsheet.scrollWidth;
	},
	/**
	 * scroll a row, col or cell to visible
	 * @param {Object} cell cell ctrl
	 */
	scrollToVisible: function (row, col, cell) {
		if (!cell) return;
		
		var sheet = this.sheet;
		if (cell) {
			row = cell.r;
		    col = cell.c;
		} else {
			if (zkS.t(row) && zkS.t(col))
				cell = sheet.getCell(row, col);
			else if(zkS.t(row))
				cell = sheet.getCell(row, sheet.activeBlock.range.left);
			else if(zkS.t(col))
				cell = sheet.getCell(sheet.activeBlock.range.top, col);
		}
		
		var cellcmp = cell.comp,
			local = this,
			spcmp = this.comp,
			block = sheet.activeBlock,
			w = cellcmp.offsetWidth, // component width
			h = cellcmp.offsetHeight, // component height
			l = cellcmp.offsetLeft + block.comp.offsetLeft,//cell left in row + block left in datapanel = cell left in scroll panel 
			t = cellcmp.parentNode.offsetTop + block.comp.offsetTop,//Row top in block + block top in datapanel = cell top in scroll panel.
			sl = spcmp.scrollLeft,//current scroll left
			st = spcmp.scrollTop,//current scroll top
			sw = spcmp.clientWidth,//scroll panel width(no scroll bar)
			sh = spcmp.clientHeight,//scroll panel height(no scroll bar)
			th = sheet.topHeight, //sheet top height
			lw = sheet.leftWidth, //sheet left width
			lsl = sl, //final scroll left
			lst = st, //final scroll top
			dirty = false,
			fzr = sheet.frozenRow, //processing for frozen
			fzc = sheet.frozenCol,
			fh = 0,
			fw = 0;
		
		/**
		 * I wan to provide a feature, is moveFocus on frozenpanel then don't scroll datapanel 
		 * but if scroll to visible after a jump(for example, at the last column then click HOME)
		 * then the datapanel wil be recreate, if i dont scroll to the certain cell, 
		 * then i will see a blank (because the cell in this position is all cleared. 
		 */
		var moveonfr = moveonfc = false;
		 
		if (zkS.t(col) && fzc > -1) {
			if (fzc < col)
				fw = sheet.custColWidth.getStartPixel(fzc + 1);
			else
				moveonfc = true;// cell on frozen col, no need to scroll datapanel 
		}
	
		if (zkS.t(row) && fzr > -1) {
			if (fzr < row)
				fh = sheet.custRowHeight.getStartPixel(fzr + 1);
			else
				moveonfr = true;// cell on frozen col, no need to scroll datapanel
		}
		//end procesing for frozen
		
		/**
		 * if cell width large then scroll panel width
		 * or cell left small then scroll panel left
		 * then scroll left = cell-left
		 * 
		 * if scroll right < cell right
		 * then scroll left = cell right - scroll width (show cell right at the end of right side 
		 */
		if (zkS.t(col) && !moveonfc) {
			//cellcmp width large then scroll pane || cell-left  small then scroll=panel-left
			if (( (sw - lw) < (w - fw)) || (l-lw) < (sl+fw) ) {  
				lsl = l-lw - fw;
				dirty = true;
			 } else if(sl + sw < l + w) {
				lsl = l + w - sw;
				dirty = true;
			}
		}
		if (zkS.t(row) && !moveonfr) {
			if(((sh - th) < (h - fh)) || (t-th) < (st+fh)){ 
				lst = t-th - fh;
				dirty = true;
			} else if(st + sh < t + h) {//top large then scroll panel bottom
				lst = t+h - sh;
				dirty = true;
			}
		}

		if (dirty) {
			this._visFlg = true;
			spcmp.scrollLeft = lsl;
			spcmp.scrollTop = lst;
			//after this , onScroll will be fired.
		}
	}
});