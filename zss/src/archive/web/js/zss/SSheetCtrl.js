/* SSheetCtrl.js

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
	_skey = [32,106,107,109,110,111,186,187,188,189,190,191,192,220,221,222,219];
	function asciiChar (charcode) {
		if ((charcode != 13 && charcode < 32) || charcode > 127) return null;
		return String.fromCharCode(charcode);
	}
	function isAsciiCharkey (keycode) {
		//48-57 number, 65-90 alpha, 
		//number pad
		//0-9 96-105
		//* 106 + 107 - 109 . 110  / 111
		//special
		//;: 186  =+ 187 ,< 188  -_ 189   .> 190  /? 191  `~ 192 
		//\| 220  } 221 '" 222 [{ 219      
		var r = ((keycode >= 48 && keycode <= 57) ||
			(keycode >= 65 && keycode <= 90) || (keycode >= 96 && keycode <= 105)),
			i = _skey.length;
		if(r) return true;
		while (i--)
			if(keycode == _skey[i]) return true;
		//firefox fire + and ; with keycode 61 & 59
		if(zk.gecko && (keycode == 61 || keycode == 59)) return true;
		
		return false;
	}
	
	function _isEvtButton (evt, flag) {
		var r = false;
		if (flag.indexOf("l") != -1)
			r |= ((evt.which) && (evt.which == 1));
		if(!r && flag.indexOf("r") != -1)
			r |= ((evt.which) && (evt.which == 3));
		if(!r && flag.indexOf("m") != -1)
			r |= ((evt.which) && (evt.which == 2));
		return r;
	}
	
	function _isLeftMouseEvt (evt) {
		return (evt.which) && (evt.which == 1);
	}
	
	function _isMiddleMouseEvt (evt) {
		return (evt.which) && (evt.which == 2);
	}
	
	function _isRightMouseEvt (evt) {
		return (evt.which) && (evt.which == 3);
	}
	function _findHeader(headers, index) {
		for(var j = 0, len = headers.length; j < len; ++j)
			if (headers[j].index == index)
				return headers[j];
		return null;
	}
	
/**
 *  SSheetCtrl controls spreadsheet
 */
zss.SSheetCtrl = zk.$extends(zk.Object, {
	$init: function (cmp, wgt) {
		this.$supers('$init', arguments);
		this._wgt = wgt;
		this.id = cmp.id;
		this.sheetid = cmp.id;//must keep this for search current
		cmp.ctrl = this;
		this.comp = cmp;
		var local = this;
		this.activeBlock = null;
		this.pageKeySize = 100;
		this._initiated = false;
		var initparm = this.initparm = {};//initial time parameter
		
		this.afterInit(function () {
			zss.Spreadsheet.initLaterAfterCssReady(local);
		}, false);
		
		//init function later queue, the function in this queue will be invoke in ZK InitialLater
		//I create this queue because ZK initial later doesn't support parameter.
		this._initLaterQ = [];//after init function queue
		this._initLaterQ.urgent = 0;//after init function queue
		
		this.state = zss.SSheetCtrl.NOFOCUS;
		
		//current server sheet index
		this.serverSheetId = wgt.getSheetId();
		
		//initial default size;
		this.topHeightDt = this.leftWidthDt = this.rowHeightDt = 
			this.colWidthDt = this.lineHeightDt = this.cellPadDt =  false;
		
		var maxCols = wgt.getMaxColumn(),
			maxRows = wgt.getMaxRow(),
			topHeight = wgt.getTopPanelHeight(),
			leftWidth = wgt.getLeftPanelWidth(),
			cellPad = wgt.getCellPadding(),
			rowHeight = wgt.getRowHeight(),
			colWidth = wgt.getColumnWidth(),
			lineHeight = wgt.getLineh();

		this.maxCols = maxCols != null ? maxCols : 10;//255;
		this.maxRows = maxRows != null ? maxRows : 10;//1000; 
		this.topHeight = topHeight != null ? topHeight : 20;
		this.leftWidth = leftWidth != null ? leftWidth : 28;
		this.cellPad = cellPad != null ? cellPad : 2;
		this.rowHeight = rowHeight != null ? rowHeight : 20;
		this.colWidth = colWidth != null ? colWidth : 80;
		this.lineHeight = lineHeight != null ? lineHeight : 20;
		var fs = wgt.getFocusRect();
		fs = fs.split(",");
		initparm.focus = new zss.Pos(zk.parseInt(fs[1]), zk.parseInt(fs[0]));//[row,col]
		
		var sel = wgt.getSelectionRect();
		sel = sel.split(",");
		initparm.selrange = new zss.Range(zk.parseInt(sel[0]), zk.parseInt(sel[1]),zk.parseInt(sel[2]),zk.parseInt(sel[3]));
		
		var hl = wgt.getHighLightRect();
		if (hl) {
			hl = hl.split(",");
			initparm.hlrange = new zss.Range(zk.parseInt(hl[0]), zk.parseInt(hl[1]), zk.parseInt(hl[2]), zk.parseInt(hl[3]));
			
			this.addSSInitLater(function() {
				var range = local.initparm.hlrange;
				local.hlArea.show = true;
				local.moveHighlight(range.left, range.top, range.right, range.bottom);
			});
			
		} else
			initparm.hlrange = new zss.Range(-1, -1, -1, -1);

		this.config = new zss.Configuration();
		
		//customized width Top Header array
		var csc = wgt.getCsc(),
			array = [];
		if (csc && csc != "") {
			csc = csc.split(",");
			var size  = csc.length;
			for (var i = 0; i < size; i = i + 4)
				array.push([zk.parseInt(csc[i]), zk.parseInt(csc[i + 1]), zk.parseInt(csc[i + 2]), 'true' == csc[i+3]]);
		}
		this.custColWidth = new zss.PositionHelper(this.colWidth, array);
		this.custColWidth.ids = new zss.Id(0, 2);
		
		//customized height Left Header array
		var csr = wgt.getCsr("csr"),
			array = [];
		if (csr && csr != "") {
			csr = csr.split(",");
			var size  = csr.length;
			for (var i = 0; i < size; i = i + 4)
				array.push([zk.parseInt(csr[i]), zk.parseInt(csr[i + 1]), zk.parseInt(csr[i + 2]), 'true' == csc[i+3]]);
		}
		this.custRowHeight = new zss.PositionHelper(this.rowHeight, array);
		this.custRowHeight.ids = new zss.Id(0, 2);

		//merge range
		var mers = wgt.getMergeRange(),
			array = [];
		if (mers && mers != "") {
			mers = mers.split(";");
			var size  = mers.length,
				r;
			for (var i = 0; i < size; i++) {
				r = mers[i].split(",");
				var range = new zss.Range(zk.parseInt(r[0]), zk.parseInt(r[1]), zk.parseInt(r[2]), zk.parseInt(r[3]));
				range.id = zk.parseInt(r[4]);
				array.push(range);
			}
		}

		//frozen row & column
		var fzr = wgt.getRowFreeze(),
			fzc = wgt.getColumnFreeze();
		this.frozenRow = fzr != null ? zk.parseInt(fzr) : -1;
		this.frozenCol = fzc != null ? zk.parseInt(fzc) : -1;
		
		this.mergeMatrix = new zss.MergeMatrix(array, this);
		
		//init inner components
		zss.SSheetCtrl._initInnerComp(this);
		
		this.innerClicking = 0;// mouse down counter to check that is focus rellay lost.

		
		this.insertSSInitLater(function () {
			local._fixSize();//fix size and scroll after initial.
		});
		this.addSSInitLater(function() {
			delete local.initparm;
		});
	},
	addEditorFocus : function(name, color){
		var x = this.focusmarkcmp,
			div = x.cloneNode(true);
		div.style.borderColor = color;
		div.style.borderWidth = "3px";
		x.parentNode.appendChild(div);
		if(!this.editorFocusMark)
			this.editorFocusMark = new Object();
		this.editorFocusMark[name] = new zss.FocusMarkCtrl(this, div, new zss.Pos(0, 0));
	},
	removeEditorFocus : function(name){
		var ctrl = this.editorFocusMark[name];
		if (ctrl) {
			ctrl.comp.parentNode.removeChild(ctrl.comp);
			ctrl.cleanup();
		}
		this.editorFocusMark[name] = null;
	},
	moveEditorFocus : function(name, color, row, col){
		if(!this.editorFocusMark || !this.editorFocusMark[name]){
			this.addEditorFocus(name, color);
		}
		this.editorFocusMark[name].relocate(row, col);
		this.editorFocusMark[name].showMark();
	},
	_resize: function () {
		if (this.invalid) return;
		this._fixSize();
		this.activeBlock.loadForVisible();
	},
	cleanup: function () {
		this.animateHighlight(false);
		this.invalid = true;

		if(this.comp) this.comp.ctrl = null;
		this.comp = this.busycmp = this.maskcmp = this.spcmp = this.topcmp = 
		this.leftcmp = this.dpcmp = this.sinfocmp = this.infocmp = this.focusmarkcmp =
		this.selareacmp = this.selchgcmp = this.hlcmp = this.wpcmp = null;
		
		if (this.dragging) {
			this.dragging.cleanup();
			this.dragging = null;
		}
		
		this.sp.cleanup();
		this.dp.cleanup();
		this.tp.cleanup();
		this.lp.cleanup();
		this.cp.cleanup();
		this.sinfo.cleanup();
		this.info.cleanup();
		this.editor.cleanup();
		this.focusMark.cleanup();
		this.selArea.cleanup();
		this.selChgArea.cleanup();
		this.hlArea.cleanup();
		this.sp = this.dp = this.tp = this.lp = this.cp = this.sinfo = this.info = 
		this.editor = this.focusMark = this.selArea = this.selChgArea = this.hlArea = null;
		
		if (this.activeBlock)
			this.activeBlock.cleanup();
		
		this.activeBlock = this.custTHSize = this.custLHSize = this._initLaterQ =
		this._lastmdelm = this._lastmdstr = null;
	},
	/**
	 * Returns whether is dragging or not
	 * @return boolean
	 */
	isDragging: function () {
		return this.dragging ? true : false;
	},
	/**
	 * Returns whether is stop dragging or not
	 * @return boolean
	 */
	stopDragging: function () {
		if (this.dragging) {
			this.dragging.cleanup();
			this.dragging = null;
		}
	},
	/**
	 * Sets dragging status
	 * @param boolean dragging
	 */
	setDragging: function (dragging) {
		this.stopDragging();
		this.dragging = dragging;
	},
	/**
	 * Returns whether is in asynchronous state or not
	 * @return boolean
	 */
	isAsync: function(){
		return (this.state&1 == 1);
	},
	/**
	 * Add init function to queue
	 * @param fn
	 * @param arg0
	 * @param arg1
	 * @param arg2
	 * @param arg3
	 */
	addSSInitLater: function (fn, arg0, arg1, arg2, arg3) {
		this._initLaterQ.push([fn, arg0, arg1, arg2, arg3]);
	},
	/**
	 * Insert init function to the queue
	 * @param fn
	 * @param arg0
	 * @param arg1
	 * @param arg2
	 * @param arg3
	 */
	insertSSInitLater: function(fn, arg0, arg1, arg2, arg3){
		this._initLaterQ.unshift([fn,arg0,arg1,arg2,arg3]);
		this._initLaterQ.urgent ++;
	},
	_doSSInitLater: function () {
		if(this.invalid) return;
		
		var local = this,
			queu = local._initLaterQ;

		if (queu.length == 0) return;
		var urgent = queu.urgent,
			parm,
			count = 0;
		while ((parm = queu.shift())) {
			parm[0](parm[1], parm[2], parm[3], parm[4]);//fn, arg0,arg1,arg2
			if(count > urgent &&  count >= 25 ){//break a while.
				setTimeout(function(){
					local._doSSInitLater();
				}, 1);
				break;
			}
			count++;
		}
		queu.urgent = 0;
	},
	_cmdCellUpdate: function (result) {
		var type = result.type,
			row = result.r,
			col = result.c,
			value = result.val,
			server = result.server; //use editValue from server
		switch(type){
		case "udcell":
			this._updateCell(result);
			break;
		case "startedit":
			var dp = this.dp;
			if (!this.dp._moveFocus(row, col, true, true)) {
				//TODO, if cell not initial, i should skip or put to delay batch? 
				break;
			}
			dp._startEditing(value, server);
			break;
		case "stopedit":
			var dp = this.dp;
			dp._stopEditing(value);
			break;
		case "canceledit":
			var dp = this.dp;
			dp.cancelEditing();
			break;
		}
	},
	_cmdBlockUpdate: function (result) {
		if (result.type == "neighbor") {//move to a neighbor block
			this.activeBlock._createCell(result);
			if (zk.ie) {
				//ie have some display error(cell overlap) when scroll up(neighbor north)
				//same issue when scroll right
				var dp = this.dp.comp,
					l = this.lp.comp;
				jq(dp).css('display', 'none');//for speed up
				jq(l).css('display', 'none');//for speed up
				
				zk(dp).redoCSS();
				zk(l).redoCSS();
				
				jq(dp).css('display', '');
				jq(l).css('display', '');
			}
		} else if (result.type == "jump") {//jump to another bolck, not a neighbor

			var oldBlock = this.activeBlock;
			this.activeBlock = zss.MainBlockCtrl.createComp(this, result.left, result.top);
			this.activeBlock._createCell(result);

			// to avoid row slip effect when replace child, i hide it, the show it (but it cause flash effect)
			//this.activeBlock.hide();
			this.dp.comp.replaceChild(this.activeBlock.comp, oldBlock.comp);
			//this.activeBlock.show();
			
			this.dp._fixSize(this.activeBlock);
			oldBlock.cleanup();
		} else if (result.type == "prune") {//jump to another bolck, not a neighbor
			
			this.activeBlock.pruneCell(result.dir, null, result.reserve);
		} else if (result.type == "ack") {//fetch cell will empty return,(not thing need to fetch)
			
		} else if (result.type == "error") {//fetch cell with exception
			
		}

		this.showMask(false);
		//close busy after all initial after load block
		//this.addSSInitLater(function(sheet){sheet.showBusy(false);},this);
		
		//bug 1951423 IE : row is broken when scroll down, st time to do ss initiallater
		var self = this;
		setTimeout(function(){
			self._doSSInitLater();//after creating cell need to invoke init later
		}, 0);
	},
	_cmdInsertRC: function (result) {
		var local = this;
		if (result.type == "column") {
			var col = result.col,
				size = result.size;
			this._insertNewColumn(col, size, result.extnm);
			//update positionHelper
			this.custColWidth.shiftMeta(col, size);
			// adjust data panel size;
			var dp = this.dp;
			dp.updateWidth(this.colWidth * size);
			
			//update maxCol
			this.maxCols = result.maxcol;

			//fix frozenCol size
			var fzc = this.frozenCol = result.colfreeze;
			if (fzc > -1) {
				this.lp._fixSize();
				this.cp._fixSize();
			}
			var block = this.activeBlock;
			if (col < block.range.left)// insert before current block, then jump
				block.reloadBlock("east");
			else 
				block._processTextOverflow(col);
		} else if (result.type == "row") {//jump to another bolck, not a neighbor
			var row = result.row,
				size = result.size;
			this._insertNewRow(row, size, result.extnm);
			//update positionHelper
			this.custRowHeight.shiftMeta(row, size);
			// adjust datapanel size;
			var dp = this.dp;
			dp.updateHeight(this.rowHeight * size);
			
			//update maxRow
			this.maxRows = result.maxrow;

			//fix frozen size
			var fzr = this.frozenRow = result.rowfreeze;;
			if (fzr > -1) {
				this.tp._fixSize();
				this.cp._fixSize();
			}
			var block = this.activeBlock;
			if (row < block.range.top)// insert before current block, then jump
				block.reloadBlock("south");
		}		
		dp._fixSize(this.activeBlock);
		this._fixSize();

		this.sendSyncblock();
		
		var self = this;
		setTimeout(function () {
			self._doSSInitLater();//after creating cell need to invoke init later
		},0);
	},
	_cmdRemoveRC: function (result, shfitsize) {
		var lfv = true;
		if (result.type == "column") {
			var col = result.col,
				size = result.size;
			
			this._removeColumn(col,size,result.extnm)
			
			// adjust datapanel size;
			var dp = this.dp,
				w = this.custColWidth.getStartPixel(col);
			w = this.custColWidth.getStartPixel(col + size) - w;
			dp.updateWidth(-w);
			
			//update positionHelper
			if(shfitsize) this.custColWidth.unshiftMeta(col,size);
			
			//update maxCol
			this.maxCols = result.maxcol;

			//fix frozenCol size
			this.frozenCol = result.colfreeze;
			this.lp._fixSize();
			this.cp._fixSize();
			
			var block = this.activeBlock;
			if (col < block.range.left) {// insert before current block, then jump
				block.reloadBlock("east");
				lfv = false;
			} else
				block._processTextOverflow(col);
			
			this._syncColFocusAndSelection(col, col + size - 1);
		} else if (result.type == "row") {//jump to another bolck, not a neighbor
			var row = result.row,
				size = result.size;
			this._removeRow(row, size, result.extnm);
			
			// adjust datapanel size;
			var dp = this.dp,
				h = this.custRowHeight.getStartPixel(row);
			h = this.custRowHeight.getStartPixel(row + size) - h;
			dp.updateHeight(-h);
			
			//update positionHelper
			if(shfitsize) this.custRowHeight.unshiftMeta(row,size);
			
			//update maxRow
			this.maxRows = result.maxrow;

			//fix frozen size
			this.frozenRow = result.rowfreeze;
			this.tp._fixSize();
			this.cp._fixSize();
			
			var block = this.activeBlock;
			if (row < block.range.top) {// insert before current block, then jump
				block.reloadBlock("south");
				lfv = false;
			}
			this._syncRowFocusAndSelection(row, row + size - 1);
		}

		dp._fixSize(this.activeBlock);
		this._fixSize();		
		this.sendSyncblock();
		
		if(lfv) this.activeBlock.loadForVisible();
		
		var pos = this.getLastFocus(),
			update;
		if (pos.row >= this.maxRows) {
			pos.row = this.maxRows - 1;
			update = true;
		}
		if (pos.column >= this.maxCols) {
			pos.column = this.maxCols - 1;
			update = true;
		}
		
		if(update) dp.moveFocus(pos.row, pos.column, true, true);
		
		var self = this;
		setTimeout(function () {
			self._doSSInitLater();//after creating cell need to invoke init later
		}, 0);
	},
	_cmdMaxcolumn: function (result) {
		var maxcol = result.maxcol,
			colfreeze = result.colfreeze;
		if (maxcol > this.maxCols) {
			// adjust datapanel size;
			var dp = this.dp,
				w = this.custColWidth.getStartPixel(this.maxCols);
			w = this.custColWidth.getStartPixel(maxcol) - w;
			dp.updateWidth(w);
			
			//update maxCol
			this.maxCols = maxcol;
			
			dp._fixSize(this.activeBlock);
			this._fixSize();
			
			this.activeBlock.loadForVisible()
		} else if (maxcol < this.maxCols) {
			var result = {};
			result.type = "column";
			result.col = maxcol;
			result.size = this.maxCols - maxcol;
			result.maxcol = maxcol;
			result.colfreeze = colfreeze;
			this._cmdRemoveRC(result, false);
		}
	},
	_cmdMaxrow: function (result) {
		var maxrow = result.maxrow,
			rowfreeze = result.rowfreeze;

		if (maxrow > this.maxRows) {
			// adjust datapanel size;
			var dp = this.dp,
				h = this.custRowHeight.getStartPixel(this.maxRows);
			h = this.custRowHeight.getStartPixel(maxrow)-h;
			dp.updateHeight(h);
			
			//update maxRow
			this.maxRows = maxrow;
			
			dp._fixSize(this.activeBlock);
			this._fixSize();
			
			this.activeBlock.loadForVisible()
		} else if (maxrow < this.maxRows) {
			var result = {};
			result.type = "row";
			result.row = maxrow;
			result.size = this.maxRows - maxrow;
			result.maxrow = maxrow;
			result.rowfreeze = rowfreeze;
			this._cmdRemoveRC(result, false);
		}
	},
	_cmdMerge: function (result){
		var type = result.type;
		if (type == "remove")
			this._removeMergeRange(result);
		else if(type == "add")
			this._addMergeRange(result);
	},
	_cmdSelection: function (result) {
		var type = result.type,
			left = result.left,
			top = result.top,
			right = result.right,
			bottom = result.bottom;
		if (type == "move") {
			this.moveCellSelection(left, top, right, bottom, true);
			var ls = this.getLastSelection();//cause of merge, selection might be change, get form last
			if (ls.left != left || ls.right != right || ls.top != top || ls.bottom != bottom) {
				this.selType = zss.SelDrag.SELCELLS;
				this._sendOnCellSelection(this.selType, ls.left, ls.top, ls.right, ls.bottom);
			}
		}
	},
	_cmdCellFocus: function (result) {
		var type = result.type,
			row = result.row,
			column = result.column;
		if (type == "move") {
			this.moveCellFocus(row, column);
			var pos = this.getLastFocus();
			if (pos.row != row || pos.column != column) //update server to new focus position
				this._sendOnCellFocused(pos.row, pos.column);
		}
	},
	_cmdRetriveFocus: function (result) {
		var type = result.type,
			row = result.row,
			column = result.column;
		if (type == "moveto") {
			//sheet.dp.selectCell(row,column,true);
			this.dp.moveFocus(row, column, true, true, true);
		} else if (type == "retrive") {
			this.dp._gainFocus(true, true);
		}
	},
	_cmdSize: function (result) {
		var type = result.type;
		if (type == "column")
			this._setColumnWidth(result.column, result.width, false, true, result.hidden, result.id);
		else if(type=="row")
			this._setRowHeight(result.row, result.height, false, true, result.hidden, result.id);
	},
	_cmdHighlight: function (result) {
		var type = result.type;
		if (type == "hide") {
			this.hlArea.show = false;
			this.hideHighlight();
		} else if(type == "show") {
			this.hlArea.show = true;
			this.moveHighlight(result.left, result.top, result.right, result.bottom);
		}
	},
	_cmdGridlines: function (show) {
		this.setDisplayGridlines(show);
	},
	_shiftMouseSelection: function(evt, row, col, selType) {
		if (zkS.isEvtKey(evt, "s") && _isLeftMouseEvt(evt)) {
			var fpos = this.getLastFocus(),
				frow = fpos.row,
				fcol = fpos.column,
				left = col < fcol ? col : fcol,
				right = col < fcol ? fcol : col,
				top = row < frow ? row :  frow,
				bottom = row < frow ? frow : row;
			
			this.moveCellSelection(left, top, right, bottom, true);
			var ls = this.getLastSelection();
			this.selType = selType;
			this._sendOnCellSelection(selType, ls.left, ls.top, ls.right, ls.bottom);
			return true;
		}
		return false;
	},
	_doMousedown: function (evt) {
		this.innerClicking++;
		var sheet = this;
		setTimeout(function() {
			if (sheet.innerClicking > 0) sheet.innerClicking--;
		}, 0);

		var elm = evt.domTarget;
		if (zkS.parentByZSType(elm, "SMask"))
			return;

		if (this.state == zss.SSheetCtrl.NOFOCUS) {
			this._nfdown = true;//down on no foucs
			this.dp._gainFocus(true);
			return;
		} else
			this._nfdown = false;

		if (!_isEvtButton(evt, "lr"))
			return;
		
		this._lastmdelm = elm;//last mouse down on element
		this._lastmdstr = "";//last mouse down str;
		
		var cmp, row, col, mx, my;
		if ((cmp = zkS.parentByZSType(elm, "SCell"))) {
			var cmpofs = zk(cmp).revisedOffset();
			mx = evt.pageX;
			my = evt.pageY;
			
			var cellpos = zss.SSheetCtrl._calCellPos(sheet, mx, my, false);//calculate if over the width
			
			row = cellpos[0];
			col = cellpos[1];
			if (this._shiftMouseSelection(evt, row, col, zss.SelDrag.SELCELLS))
				return;			
			sheet.dp.moveFocus(row, col, false, true, false, true);
			this._lastmdstr = "c";

			var ls = this.getLastSelection();//cause of merge, focus might be change, get form last
			this.selType = zss.SelDrag.SELCELLS;
			this.setDragging(new zss.SelDrag(sheet, this.selType, ls.top, ls.left,
					_isLeftMouseEvt(evt) ? "l" : "r", ls.right));
			
			//start hyperlink follow up
			if(_isLeftMouseEvt(evt) && this.selArea)
				this.selArea._startHyperlink(elm);
			
		} else if ((cmp = zkS.parentByZSType(elm, "SSelDot", 1)) != null) {
			//modify selection
			if(_isLeftMouseEvt(evt)) {//TODO support right mouse down
				if (!this.selType)
					this.selType = zss.SelDrag.SELCELLS;
				var action = this.selType | zss.SelChgDrag.MODIFY;
				this.setDragging(new zss.SelChgDrag(sheet, action));
			}
		} else if ((cmp = zkS.parentByZSType(elm, ["SSelInner", "SFocus", "SHighlight"], 1)) != null) {
			//Mouse down on Selection / Focus Block
			mx = evt.pageX;
			my = evt.pageY;
			
			var cellpos = zss.SSheetCtrl._calCellPos(sheet, mx, my, false);
			row = cellpos[0];
			col = cellpos[1];
			if (this._shiftMouseSelection(evt, row, col, zss.SelDrag.SELCELLS))
				return;			
			this._lastmdstr = "c";

			if (_isLeftMouseEvt(evt) || jq(cmp).attr('zs.t') == "SHighlight") {
				sheet.dp.moveFocus(row, col, false, true, false, true);
				var ls = this.getLastSelection();//cause of merge, focus might be change, get form last
				this.selType = zss.SelDrag.SELCELLS;
				this.setDragging(new zss.SelDrag(sheet, this.selType, ls.top, ls.left,
						_isLeftMouseEvt(evt) ? "l" : "r", ls.right));
				
				//start hyperlink follow up
				if (this.selArea)
					this.selArea._startHyperlink();
			}
		} else if ((cmp = zkS.parentByZSType(elm, ["SSelect"], 1)) != null) {
			mx = evt.pageX;
			my = evt.pageY;
			var cellpos = zss.SSheetCtrl._calCellPos(sheet, mx, my, false);
			row = cellpos[0];
			col = cellpos[1];
			if (this._shiftMouseSelection(evt, row, col, zss.SelDrag.SELCELLS))
				return;			
			this._lastmdstr = "c";
			
			if(_isLeftMouseEvt(evt)){//TODO support right mouse down
				range = sheet.selArea.lastRange;
				
				//adjust row,col to selection range
				if (row > range.bottom)
					row = range.bottom;
				else if (row < range.top)
					row = range.top;

				if (col > range.right)
					col = range.right;
				else if(col < range.left)
					col = range.left;
			
				if (!this.selType)
					this.selType = zss.SelDrag.SELCELLS;
				var action = this.selType | zss.SelChgDrag.MOVE; 
				this.setDragging(new zss.SelChgDrag(sheet, action, row, col));
				
				//start hyperlink follow up
				if (this.selArea)
					this.selArea._startHyperlink();
			}
		} else if ((cmp = zkS.parentByZSType(elm, "SLheader")) != null 
			|| (cmp = zkS.parentByZSType(elm, "STheader")) != null) {
			var type = (jq(cmp).attr('zs.t') == "SLheader") ? zss.Header.VER : zss.Header.HOR,
				row, col, onsel,	//process select row or column
				ls = this.selArea.lastRange;

			this._lastmdstr = "h";
			if (type == zss.Header.HOR) {
				row = -1;
				col = cmp.ctrl.index;
				if(col >= ls.left && col <= ls.right) {
					onsel = true;
				}
			} else {
				row = cmp.ctrl.index;
				col = -1;
				if(row >= ls.top && row <= ls.bottom) {
					onsel = true;
				}
			}

			if (_isLeftMouseEvt(evt) || !onsel) {
				var range = zss.SSheetCtrl._getVisibleRange(this),
					seltype;
				if (row == -1) {//column
					if (zkS.isEvtKey(evt, "s") && _isLeftMouseEvt(evt)) {
						var fpos = this.getLastFocus(),
							fcol = fpos.column;
						sheet.moveColumnSelection(col, fcol);
						var ls = this.getLastSelection();
						this.selType = zss.SelDrag.SELCOL;
						this._sendOnCellSelection(this.selType, ls.left, ls.top, ls.right, ls.bottom);
						return;
					}
					var fzr = sheet.frozenRow;
					sheet.dp.moveFocus((fzr > -1 ? 0 : range.top), col, true, true, false, true);
					//sheet.dp.selectCell((fzr > -1 ? 0 : range.top), col, true);//force move to first visible cell or 0 if frozenRow
					sheet.moveColumnSelection(col);
					seltype = zss.SelDrag.SELCOL;
				} else {
					if (zkS.isEvtKey(evt, "s") && _isLeftMouseEvt(evt)) {
						var fpos = this.getLastFocus(),
							frow = fpos.row;
						sheet.moveRowSelection(row, frow);
						var ls = this.getLastSelection();
						this.selType = zss.SelDrag.SELROW;
						this._sendOnCellSelection(this.selType, ls.left, ls.top, ls.right, ls.bottom);
						return;
					}
					var fzc = sheet.frozenCol;
					sheet.dp.moveFocus(row, (fzc > -1 ? 0 : range.left), true, true, false, true);
					//sheet.dp.selectCell(row, (fzc > -1 ? 0 : range.left),true);//force move to first visible cell
					sheet.moveRowSelection(row);
					seltype = zss.SelDrag.SELROW;
				}
				sheet.selType = seltype;
				this.setDragging(new zss.SelDrag(sheet, seltype, row, col, _isLeftMouseEvt(evt) ? "l" : "r"));
			}
		} else if ((cmp = zkS.parentByZSType(elm, "SCorner", 1)) != null) {
			var ls = this.getLastSelection(),
				left = 0,
				top = 0,
				right = this.maxCols - 1,
				bottom = this.maxRows - 1;
			
			if (left != ls.left || top != ls.top || right != ls.right || bottom != ls.bottom) {
				this.moveCellSelection(left, top, right, bottom);
				this.selType = zss.SelDrag.SELALL;
				this.setDragging(new zss.SelDrag(sheet, this.selType, 0, 0, _isLeftMouseEvt(evt) ? "l" : "r"));
			}
		}
		this._lastmdstr = this._lastmdstr + "_" + row + "_" + col;
	},
	_doMouseup: function (evt) {
		if (this.isAsync())//wait async event, skip mouse click;
			return;

		//bug#1974069, leftkey & has last mouse down element,
		if (_isLeftMouseEvt(evt) && this._lastmdelm && zkS.parentByZSType(this._lastmdelm, zk.ie8 ? ["SCell", "SHighlight", "SSelInner"] : ["SCell", "SHighlight"], 1) != null) {
			this._doMouseclick(evt, "lc", this._lastmdelm);
		}
		this._lastmdelm = null;
	},
	_doMouseleftclick: function (evt) {
		if(this.isAsync())//wait async event, skip mouse click;
			return;

		this._doMouseclick(evt, "lc");
		evt.stop();
	},
	_doMouserightclick : function (evt) {
		if (this.isAsync())//wait async event, skip mouse click;
			return;
		this._doMouseclick(evt, "rc");
		evt.stop();//always stop right (context) click.
	},
	_doMousedblclick: function (evt) {
		if (this.isAsync())//wait async event, skip mouse click;
			return;
		this._doMouseclick(evt, "dbc");
		evt.stop();
	},
	/**
	 * @param zk.Event, mouse event
	 * @param string type "lc" for left click, "rc" for right click, "dbc" for double click
	 */
	_doMouseclick: function (evt, type, element) {
		if (this._nfdown) {
			if (this.selArea) 
				this.selArea._stopHyperlink();
			return; // don't care click if it was fired when nofocus mouse down
		}
		var sheet = this,
			wgt = sheet._wgt,
			elm = (element) ? element : evt.domTarget,
			cmp,
			mx = my = 0,//mouse offset, against body 
			shx = shy = 0,//mouse offset against sheet
			row, 
			col,
			md1 = zkS._getMouseData(evt, this.comp),
			mdstr = "";
		//Click on Cell
		if ((cmp = zkS.parentByZSType(elm, "SCell", 1)) != null) {
			var cellcmp = cmp,
				sheetofs = zk(sheet.comp).revisedOffset(),
			//TODO there is a bug in opera, when a cell is overflow, zk.revisedOffset can get correct component offset 
				cmpofs = zk(cellcmp).revisedOffset();
			
			mx = evt.pageX;
			my = evt.pageY;
			shx = Math.round(mx - sheetofs[0]);
			shy = Math.round(my - sheetofs[1]);
			
			var x = mx - cmpofs[0],
				cellpos = zss.SSheetCtrl._calCellPos(sheet, mx, my, false);
			row = cellpos[0];
			col = cellpos[1];
			mdstr = "c_" + row + "_" + col;

			if (this._lastmdstr == mdstr)
				wgt.fireCellEvt(type, shx, shy, md1[2], row, col, mx, my);

			if (type == 'lc' && this.selArea) {
				this.selArea._setHyperlinkElment(elm);
				this.selArea._tryAndEndHyperlink(row, col, evt);
			}
		} else if((cmp = zkS.parentByZSType(elm, "SSelDot", 1)) != null) {
		//TODO
		} else if((cmp = zkS.parentByZSType(elm, [zk.ie8 ? "SSelInner" : "SSelect", "SFocus", "SHighlight"], 1)) != null ) {
			//Mouse click on Selection / Focus Block
			var sheetofs = zk(sheet.comp).revisedOffset();
			mx = evt.pageX;
			my = evt.pageY;
			shx = Math.round(mx - sheetofs[0]);
			shy = Math.round(my - sheetofs[1]);
			
			var cellpos = zss.SSheetCtrl._calCellPos(sheet, mx, my, false),
				cx = cellpos[4]; //x relative to cell 
			row = cellpos[0];
			col = cellpos[1];
			
			//try hyperlink
			if (this.selArea) {
				this.selArea._tryAndEndHyperlink(row, col, evt);
			}
			
			mdstr = "c_" + row + "_" + col;
			if (this._lastmdstr == mdstr) {
				wgt.fireCellEvt(type, shx, shy, md1[2], row, col, mx, my);
				if (type == "dbc") {
					sheet.dp.startEditing();
					sheet.dp._openEditbox();
				}
			}
		} else if ((cmp = zkS.parentByZSType(elm, "STheader",1)) != null ||
			(cmp = zkS.parentByZSType(elm, "SLheader",1)) != null) {
			//Click on header
			var headercmp = cmp,
				sheetofs = zk(sheet.comp).revisedOffset();
			
			mx = evt.pageX;
			my = evt.pageY;
			shx = Math.round(mx - sheetofs[0]);
			shy = Math.round(my - sheetofs[1]);
			
			if (headercmp.ctrl.type == zss.Header.HOR) {
				row = -1;
				col = headercmp.ctrl.index;
			} else {
				row = headercmp.ctrl.index;
				col = -1;
			}
			mdstr = "h_" + row + "_" + col;
			if (this._lastmdstr == mdstr) {
				wgt.fireHeaderEvt(type, shx, shy, md1[2], row, col, mx, my);
			}
		}
	},
	_sendOnCellFocused: function (row, col) {
		var wgt = this._wgt;
		wgt.fire('onZSSCellFocused', {sheetId: this.serverSheetId, row: row, col : col}, wgt.isListen('onCellFocused') ? {toServer: true} : null);
	},
	_sendOnCellSelection: function (type, left, top, right, bottom) {
		this._wgt.fire('onCellSelection',
				{sheetId: this.serverSheetId, action: type, left: left, top: top, right: right, bottom: bottom});
	},
	_sendOnSelectionChange: function (action, left, top, right, bottom, orgleft, orgtop, orgright, orgbottom) {
		this._wgt.fire('onSelectionChange',
				{sheetId: this.serverSheetId, action: action, left: left,top: top, right: right, bottom: bottom, orgileft: orgleft, orgitop: orgtop, orgiright: orgright, orgibottom: orgbottom});
	},
	_sendOnHyperlink: function (row, col, href, type, evt) {
		var wgt = this._wgt,
			data = zk.copy(evt.data, {sheetId: this.serverSheetId, row: row, col: col, href: href, type: type});
		wgt.fire('onHyperlink', data, wgt.isListen('onHyperlink') ? {toServer: true} : null);
	},
	_timeoutId: null,
	_fireOnOpenAndEdit: function (time) { //open Editbox and start editing
		clearTimeout(this._timeoutId);
		this._timeoutId = setTimeout(this.proxy(this._onOpenAndEdit), time >= 0 ? time : 100);
	},
	_onOpenAndEdit: function () {
		this.dp._startEditing(); //open client side edit box to catch the user input
	},
	_doKeypress: function (evt) {
		if (this._skipress)//wait async event, skip
			return;
		
		var charcode = evt.which,
			c = asciiChar(charcode);

		//ascii, not editing, not special key
		if (c != null && !(evt.altKey || evt.ctrlKey) && this.state != zss.SSheetCtrl.EDITING) {
			if (this.state == zss.SSheetCtrl.START_EDIT) //startEditing but not get response from server yet
				this._clienttxt += c; //this._clienttxt is cleared in DataPanel#_startEditing()
			else if (this.state == zss.SSheetCtrl.FOCUSED) {
				if (this.dp.startEditing(evt, c)) //fire to server so user can override the result
					this._clienttxt = c;
			}
			//bug #117: Barcode Scanner data incomplete
			this._fireOnOpenAndEdit(); 				
			evt.stop();
		}
	},
	_doKeydown: function(evt) {
		this._skipress = false;
		//wait async event, skip
		//handle spreadsheet common keydown event

		if (this.isAsync()) return;
		
		var keycode = evt.keyCode,
			ctrl;
		switch (keycode) {
		case 33: //PgUp
			this.dp.movePageup(evt);
			evt.stop();
			break;
		case 34: //PgDn
			this.dp.movePagedown(evt);
			evt.stop();
			break;
		case 35: //End
			if(this.state != zss.SSheetCtrl.FOCUSED) break; //editing
			this.dp.moveEnd(evt);
			evt.stop();
			break;
		case 36: //Home
			if(this.state != zss.SSheetCtrl.FOCUSED) break;//editing
			this.dp.moveHome(evt);
			evt.stop();
			break;
		case 37: //Left
			if(this.state != zss.SSheetCtrl.FOCUSED) break;//editing
			this.dp.moveLeft(evt);
			evt.stop();
			break;
		case 38: //Up
			if(this.state != zss.SSheetCtrl.FOCUSED) break;//editing
			this.dp.moveUp(evt);
			evt.stop();
			break;
		case 9://tab;
			if (this.state == zss.SSheetCtrl.EDITING){
				if (evt.altKey || evt.ctrlKey)
					break;

				this.dp.stopEditing(evt.shiftKey ? "moveleft" : "moveright");//invoke move right after stopEdit
				evt.stop();
				break;
			}
			if (evt.shiftKey) {
				this.dp.moveLeft();
				evt.stop();
			} else if (!(evt.altKey || evt.ctrlKey)) {
				this.dp.moveRight();
				evt.stop();
			}
			break;
		case 39: //Right
			if (this.state != zss.SSheetCtrl.FOCUSED) break;//editing
			this.dp.moveRight(evt);
			evt.stop();
			break;
		case 40: //Down
			if (this.state != zss.SSheetCtrl.FOCUSED) break;//editing
			this.dp.moveDown(evt);
			evt.stop();
			break;
		case 113: //F2
			if(this.state == zss.SSheetCtrl.FOCUSED) {
				this.dp.startEditing(evt);
				this.dp._openEditbox();
			}
			evt.stop();
			break;
		case 13://Enter
			if (this.state == zss.SSheetCtrl.EDITING){
				if(evt.altKey || evt.ctrlKey){
					this.editor.newLine();
					evt.stop();
					break;
				}
				this.dp.stopEditing("movedown");//invoke move down after stopEdit
				evt.stop();
			} else if (this.state == zss.SSheetCtrl.FOCUSED) {
				if (!this._wgt._copysrc) {
					this.dp.moveDown(evt);
					evt.stop();
				}
			}
			break;
		case 27://ESC
			if (this.state == zss.SSheetCtrl.EDITING) {
				this.dp.cancelEditing(evt);
				evt.stop();
			} else if(this.state == zss.SSheetCtrl.FOCUSED) {
				//TODO should i send onCancel here?
			}
			break;
		}
		//in my notebook,some keycode ex : LEFT(37) and RIGHT(39) will fire keypress after keydown,
		//it confuse with the ascii value "%' and ''', so add this to do some controll in key press
		if (!isAsciiCharkey(keycode))
			this._skipress = true;
	},
	_doKeyup: function(evt) {
		if(this._skipress) delete this._skipress;
	},
	/**
	 * resize Sheet
	 * @param {String} w width string
	 * @param {String} h height string
	 */
	resizeTo: function(w , h) {
		//don't use style-class, use style, because style of sheet is the colsest style. 
		var sheetcmp = this.comp;
		if (w)
			jq(sheetcmp).css('width', w);
		if (h)
			jq(sheetcmp).css('height', h);
		
		var local = this;
		setTimeout(function(){
			local._resize();
		}, 0);
	},
	/**fix sp, tp and lp size*/
	_fixSize: function () {
		var sheetcmp = this.comp,
			spcmp = this.sp.comp,
			w = zk(sheetcmp).offsetWidth() - 2,//2 is border width
			h = zk(sheetcmp).offsetHeight() - 2;//2 is border width
		if (h <= 0)
			//if user doesn't set the height of style sheet set height on it's parent, 
			//then we will get a zero height, so , i assign a default height here
			h = 100;

		var barHeight = zkS._hasScrollBar(spcmp) ? zss.Spreadsheet.scrollWidth : 0,
			barWidth = zkS._hasScrollBar(spcmp, true) ? zss.Spreadsheet.scrollWidth : 0,
			tw = w - this.leftWidth- barWidth,
			lh = h - this.topHeight - barHeight,
			dpcomp = this.dp.comp,
			zkdp = zk(dpcomp), //bug #61: Fronzen row/column does not comply with Spreadsheet's maxrows/maxcolumn
			dpw = zkdp.offsetWidth() - this.leftWidth,
			dph = zkdp.offsetHeight() - this.topHeight,
			rw = Math.min(dpw, tw),
			rh = Math.min(dph, lh);
		
		this.tp._updateWidth(rw);
		this.lp._updateHeight(rh);
		this.sp._doScrolling();
	},
	/**
	 * Returns last focus position
	 * @return zss.Pos
	 */
	getLastFocus: function () {
		return new zss.Pos(this.focusMark.row, this.focusMark.column);
	},
	/**
	 * Returns last selection range
	 * @return zss.Range
	 */
	getLastSelection: function () {
		var range = this.selArea.lastRange;
		return !range ? null : new zss.Range(range.left, range.top, range.right, range.bottom);
	},
	/**
	 * Sets column selection
	 * @param int from
	 * @param int to
	 */
	moveColumnSelection: function (from, to) {
		if (!to && to != 0)
			to = from;

		var t = from;
		if (from > to) {
			from = to;
			to = t;
		}
		this.moveCellSelection(from, 0, to, this.maxRows - 1);
	},
	/**
	 * Sets row selection
	 * @param int from 
	 * @param int to
	 */
	moveRowSelection: function (from, to) {
		if(!to && to!=0){
			to = from;
		}
		var t = from;
		if (from > to) {
			from = to;
			to = t;
		}
		this.moveCellSelection(0, from, this.maxCols - 1, to);
	},
	/**
	 * set column width
	 * @param {Object, int} col column index or header of column component
	 * @param {Object} width the new width
	 */
	setColumnWidth: function (col, width) {
		this._setColumnWidth(col, width, true, true);
	},
	_setColumnsWidth: function (rng, width, fireevent, loadvis, hidden, metaid) {
		var left = rng.left,
			right = rng.right;
		for(var col=left; col<=right; ++col)
			this._setColumnWidth(col, width, fireevent, loadvis, hidden, metaid);
	},
	_setColumnWidth: function (col, width, fireevent, loadvis, hidden, metaid) {
		var sheetid = this.sheetid,
			custColWidth = this.custColWidth,
			oldw = custColWidth.getSize(col);
		if (width < 0)
			width = 0;

		//update customized width
		var meta = custColWidth.getMeta(col),
			zsw;
		if (hidden === undefined) {
			hidden = (width == 0);
		}
		if (hidden)
			width = oldw;
		
		//adjust cell width, check also:_updateCustomDefaultStyle
		var cp = this.cellPad,
			cellwidth,
			celltextwidth = width - 2 * cp - 1;// 1 is border width//zk.revisedSize(colcmp,width);//
		
		//bug 1989680
		var fixpadding = false;
		if (celltextwidth < 0) {
			fixpadding = true;
			celltextwidth = width - 1;
		}
		cellwidth = zk.ie || zk.safari || zk.opera ? celltextwidth : width;
		
		var name = "#" + sheetid,
			createbefor = ".zs_header";
		if(zk.opera) //opera bug, it cannot insert rul to special position
			createbefor = true;

		if (!meta) {
			//append style class to column header and cell
			zsw = zkS.t(metaid) ? metaid : custColWidth.ids.next();
			custColWidth.setCustomizedSize(col, width, zsw, hidden);
			this._appendZSW(col, zsw);
		} else {
			zsw = zkS.t(metaid) ? metaid : meta[2];
			custColWidth.setCustomizedSize(col, width, zsw, hidden);
		}

		if (width <= 0 || hidden)
			zcss.setRule(name + " .zsw" + zsw, "display", "none", createbefor, sheetid + "-sheet");
		else {
			zcss.setRule(name + " .zsw" + zsw, ["display", "width"], ["", cellwidth + "px"], createbefor, sheetid + "-sheet");
			zcss.setRule(name + " .zswi" + zsw, "width", celltextwidth + "px", createbefor, sheetid + "-sheet");
			//bug 1989680
			if (fixpadding)
				zcss.setRule(name + " .zsw" + zsw, "padding", "0px", createbefor, sheetid + "-sheet");
			else
				zcss.setRule(name + " .zsw" + zsw, "padding", "", createbefor, sheetid + "-sheet");
		}

		//set merged cell width;
		var ranges = this.mergeMatrix.getRangesByColumn(col),
			size = ranges.length,
			range,
			cssid = sheetid + "-sheet" +((zk.opera) ? "-opera" : "");//opera bug, it cannot insert rul to special position
		
		for (var i = 0; i < size; i++) {
			range = ranges[i];
			var w = custColWidth.getStartPixel(range.right + 1);
			w -= custColWidth.getStartPixel(range.left);

			celltextwidth = w - 2 * cp - 1;// 1 is border width//zk.revisedSize(colcmp,width);//
			fixpadding = false;
			if (celltextwidth < 0) {
				fixpadding = true;
				celltextwidth = w - 1;
			}
			cellwidth = zk.ie || zk.safari || zk.opera ? celltextwidth : w;

			if (w < 0)
				zcss.setRule(name+" .zsmerge"+range.id,"display","none",true,cssid);
			else {
				zcss.setRule(name+" .zsmerge"+range.id,"width",cellwidth+"px",true,cssid);
				zcss.setRule(name+" .zsmerge"+range.id+" .zscelltxt","width",celltextwidth+"px",true,cssid);
				if (fixpadding)
					zcss.setRule(name+" .zsmerge"+range.id,"padding", "0px",true,cssid);
				else
					zcss.setRule(name+" .zsmerge"+range.id,"padding", "", true,cssid);
			}
		}
		
		//adjust header
		var tp = this.tp, 
			headers = tp ? tp.headers : null,
			header = headers ? _findHeader(headers, col) : null;
		if (header)
			header.setColumnHeader(hidden);
		
		if (col < this.maxCols) {
			//adjust datapanel size;
			var dp = this.dp;
			dp.updateWidth((hidden ? 0 : width) - oldw);
		
			//process text overflow when resize column
			var block = this.activeBlock;
			block._processTextOverflow(col);
			
			if (this.cp.block)
				this.cp.block._processTextOverflow(col);	

			if(this.tp.block)
				this.tp.block._processTextOverflow(col);	

			if(this.lp.block)
				this.lp.block._processTextOverflow(col);	
			
			var fzc = this.frozenCol;
			
			if (fzc >= col) {
				this.lp._fixSize();
				this.cp._fixSize();
			}

			//update datapanel padding
			dp._fixSize(this.activeBlock);
			
			if(loadvis) this.activeBlock.loadForVisible();
		
			var local = this;
			setTimeout(function(){
				local._fixSize();
			}, 0);

			this._wgt.syncWidgetPos(-1, col);
		}

		//sync focus and selection area
		this._syncColFocusAndSelection(col, col);

		if (fireevent) {
			this._wgt.fire('onZSSHeaderModif', 
					{sheetId: this.serverSheetId, type: "top", event: "size", index: col, newsize: width, id: zsw, hidden: hidden},
					{toServer: true}, 25);
		}
	},
	_syncColFocusAndSelection: function(left, right) {
		var focPos = this.getLastFocus(),
			fCol = focPos.column,
			ls = this.getLastSelection(),
			selL = ls.left,
			selR = ls.right;
		if (left <= fCol && fCol <= right)
			this.moveCellFocus(focPos.row, fCol);
		if (right >= selL && left <= selR)
			this.moveCellSelection(selL, ls.top, selR, ls.bottom);
	},
	_syncRowFocusAndSelection: function(top, bottom) {
		var focPos = this.getLastFocus(),
			fRow = focPos.row,
			ls = this.getLastSelection(),
			selT = ls.top,
			selB = ls.bottom;
		if (top <= fRow && fRow <= bottom)
			this.moveCellFocus(fRow, focPos.column);
		if (bottom >= selT && top <= selB)
			this.moveCellSelection(ls.left, selT, ls.right, selB);
	},
	_appendZSW: function(col, zsw) {
		this.activeBlock.appendZSW(col, zsw);
		this.cp.appendZSW(col, zsw);
		this.tp.appendZSW(col, zsw);
		this.lp.appendZSW(col, zsw);
	},
	_appendZSH: function(row, zsh) {
		this.activeBlock.appendZSH(row, zsh);
		this.cp.appendZSH(row, zsh);
		this.tp.appendZSH(row, zsh);
		this.lp.appendZSH(row, zsh);
	},
	/**
	 * set row height
	 * @param {Object} row row index or header of row component
	 * @param {Object} height new height of row
	 */
	setRowHeight: function(row, height) {
		this._setRowHeight(row, height, true, true);
	},
	_setRowsHeight: function(rng, height, fireevent, loadvis, hidden, metaid) {
		var top = rng.top,
			bottom = rng.bottom;
		for(var row=top; row<=bottom; ++row)
			this._setRowHeight(row, height, fireevent, loadvis, hidden, metaid);
	},
	_setRowHeight: function(row, height, fireevent, loadvis, hidden, metaid) {
		var sheetid = this.sheetid,
			custRowHeight = this.custRowHeight,
			oldh = custRowHeight.getSize(row);
		height = height <= 0 ? 0 : height;

		var name = "#" + sheetid,
			meta = custRowHeight.getMeta(row),
			zsh;
		if (hidden === undefined) {
			hidden = (height == 0);
		} 
		if (hidden)
			height = oldh;
			
		var cellheight;// = zk.revisedSize(colcmp,height,true);
		
		if(zk.ie || zk.safari || zk.opera)
			//1989680
			cellheight = height > 0 ? height - 1 : 0;
		else
			cellheight = height;

		if (!meta) {
			//append style class to column header and cell
			zsh = zkS.t(metaid) ? metaid : custRowHeight.ids.next();
			custRowHeight.setCustomizedSize(row, height, zsh, hidden);
			this._appendZSH(row, zsh);
		} else {
			zsh = zkS.t(metaid) ? metaid : meta[2];
			custRowHeight.setCustomizedSize(row, height, zsh, hidden);
		}
		
		var createbefor = ".zs_header";
		if (zk.opera)//opera bug, it cannot insert rul to special position
			createbefor = true;

		if (height <= 0 || hidden) {
			zcss.setRule(name + " .zslh" + zsh, "display", "none", createbefor, this.sheetid + "-sheet");
			zcss.setRule(name + " .zsh" + zsh, "display", "none", createbefor, this.sheetid + "-sheet");
		} else {
			zcss.setRule(name + " .zsh" + zsh, ["display", "height"],["", height + "px"], createbefor, this.sheetid+"-sheet");
			zcss.setRule(name + " .zshi" + zsh, "height", cellheight + "px", createbefor, this.sheetid + "-sheet");//both zscell and zscelltxt
			var h2 = (height > 0) ? height - 1 : 0;
			zcss.setRule(name + " .zslh" + zsh, ["display", "height", "line-height"], ["", h2 + "px", h2 + "px"], createbefor, this.sheetid + "-sheet");
		}

		//adjust header
		var lp = this.lp,
			headers = lp ? lp.headers : null,
			header = headers ? _findHeader(headers, row) : null;
		if (header)
			header.setRowHeader(hidden);
		
		if (row < this.maxRows) {
			//adjust datapanel size;
			var dp = this.dp;
			dp.updateHeight((hidden ? 0 : height) - oldh);
		
			var fzr = this.frozenRow;
			if (fzr >= row) {
				this.tp._fixSize();
				this.cp._fixSize();
			}
	
			dp._fixSize(this.activeBlock);
			
			if (loadvis) this.activeBlock.loadForVisible();
		
			var local = this;
			setTimeout(function () {
				local._fixSize();
			}, 0);

			this._wgt.syncWidgetPos(row, -1);
		}
		//sync focus and selection area
		this._syncRowFocusAndSelection(row, row);

		if (fireevent) {
			this._wgt.fire('onZSSHeaderModif', 
					{sheetId: this.serverSheetId, type: "left", event: "size", index: row, newsize: height, id: zsh, hidden: hidden},
					{toServer: true}, 25);
		}

	},
	_updateCell: function (result) {
		var row = result.r,
			col = result.c,
			cell = this.activeBlock.getCell(row, col),
			parm = {"txt": result.val,"edit":result.edit};
		zkS.copyParm(result, parm, ["st", "ist", "wrap", "hal", "rbo", "merr", "merl"]);
		
		if (cell)//update when cell exist
			zss.Cell.updateCell(cell, parm);

		if (this.cp.block) {
			cell = this.cp.block.getCell(row, col);
			if (cell) zss.Cell.updateCell(cell, parm);
		}
		if (this.tp.block) {
			cell = this.tp.block.getCell(row, col);
			if (cell) zss.Cell.updateCell(cell, parm);
		}
		if (this.lp.block) {
			cell = this.lp.block.getCell(row, col);
			if (cell) zss.Cell.updateCell(cell, parm);
		}
	},
	_updateHeaderSelectionCss: function (range, remove) {
		var top = range.top,
			bottom = range.bottom,
			left = range.left,
			right = range.right;

		//hor
		this.tp.updateSelectionCSS(left, right, remove);
		this.lp.updateSelectionCSS(top, bottom, remove);
		if (this.cp.tp)
			this.cp.tp.updateSelectionCSS(left, right, remove);	
		if (this.cp.lp)
			this.cp.lp.updateSelectionCSS(top, bottom, remove);
	},
	/**
	 * Sets whether display the gridlines for this sheet.
	 */
	setDisplayGridlines: function (show) {
		var sheetid = this.sheetid,
			name = '#' + sheetid,
			bc = show ? '':'#FFFFFF';
		zcss.setRule(name + ' .zscell', ['border-bottom-color', 'border-right-color'],[bc, bc], true, sheetid+'-sheet');
	},
	/**
	 * Sets the cell's selection area and display it
	 * 
	 * @param int left column start index
	 * @param int top row start index
	 * @param int right column end index
	 * @param int bottom row end index
	 * @param boolean snap whether snap to merge cell border
	 */
	moveCellSelection: function (left, top, right, bottom, snap) {
		var lastRange = this.selArea.lastRange;
		if (lastRange)
			this._updateHeaderSelectionCss(lastRange, true);

		var show = !(this.state == zss.SSheetCtrl.NOFOCUS);
		if (snap) {
			var maxr = right,
				minl = left;
	
			//Selection shall snap to merge area
			for (var r = bottom; r >= top; --r) {
				var cellR = this.getCell(r, maxr);
				if (cellR && cellR.merr > maxr) maxr = cellR.merr;
				var cellL = this.getCell(r, minl);
				if (cellL && cellL.merl < minl) minl = cellL.merl;
			}
			right = maxr;
			left = minl;
			//TODO when UI support vertical merge, need to handle maxb(bottom) and mint(top)
			/*
			var maxb = bottom,
				mint = top;
			for (var c = right; c >= left; --c) {
				var cellB = this.getCell(maxb, c);
				if (cellB && cellB.merb > maxb) maxb = cellB.merb;
				var cellT = this.getCell(mint, c);
				if (cellT && cellT.mert < mint) mint = cellT.mert;
			}
			bottom = maxb;
			top = mint;
			*/
		} else {
			var cell = this.getCell(top, left);
			//only show merged selection when selecing on same row
			if (top == bottom && cell && cell.merr > right)
				right = cell.merr;
		}
		
		var selRange = new zss.Range(left, top, right, bottom);
		this.selArea.relocate(selRange);
		
		if (show) {
			this._updateHeaderSelectionCss(selRange,false);
			this.selArea.showArea();
		}
		
		if (this.tp.selArea) {
			this.tp.selArea.relocate(selRange);
			if(show) this.tp.selArea.showArea();
		}
		if (this.lp.selArea) {
			this.lp.selArea.relocate(selRange);
			if(show) this.lp.selArea.showArea();
		}
		if (this.cp.selArea) {
			this.cp.selArea.relocate(selRange);
			if(show) this.cp.selArea.showArea();
		}
	},
	/**
	 * Hides cell's selection area 
	 */
	hideCellSelection: function () {
		this.selArea.hideArea();
		if (this.tp.selArea)
			this.tp.selArea.hideArea();

		if (this.lp.selArea)
			this.lp.selArea.hideArea();

		if (this.cp.selArea)
			this.cp.selArea.hideArea();

		var lastRange = this.selArea.lastRange;
		if (lastRange)
			this._updateHeaderSelectionCss(lastRange, true);
	},
	/**
	 * Sets the cell's selection change area and display it
	 * 
	 * @param int left column start index
	 * @param int top row start index
	 * @param int right column end index
	 * @param int bottom row end index
	 */
	moveSelectionChange : function (left, top, right, bottom) {	
		var selRange = new zss.Range(left, top, right, bottom);
		this.selChgArea.relocate(selRange);
		this.selChgArea.showArea();

		if (this.tp.selChgArea) {
			this.tp.selChgArea.relocate(selRange);
			this.tp.selChgArea.showArea();
		}
		if (this.lp.selChgArea) {
			this.lp.selChgArea.relocate(selRange);
			this.lp.selChgArea.showArea();
		}
		if (this.cp.selChgArea) {
			this.cp.selChgArea.relocate(selRange);
			this.cp.selChgArea.showArea();
		}
	},
	/**
	 * Hides cell's selection change area
	 */
	hideSelectionChange: function () {
		this.selChgArea.hideArea();
		if (this.tp.selChgArea)
			this.tp.selChgArea.hideArea();

		if (this.lp.selChgArea)
			this.lp.selChgArea.hideArea();

		if (this.cp.selChgArea)
			this.cp.selChgArea.hideArea();
	},
	/**
	 * Sets cell's focus mark
	 * @param int row row index
	 * @param int col column index
	 */
	moveCellFocus: function (row, col) {
		var show = !(this.state == zss.SSheetCtrl.NOFOCUS),
			cell = this.getCell(row, col);
		if (cell && cell.merl < col) //check if a merged cell
			col = cell.merl;
		this.focusMark.relocate(row, col);
		if (show)
			this.focusMark.showMark();
		if (this.tp.focusMark) {
			this.tp.focusMark.relocate(row, col);
			if(show) this.tp.focusMark.showMark();
		}
		if (this.lp.focusMark) {
			this.lp.focusMark.relocate(row, col);
			if(show) this.lp.focusMark.showMark();
		}
		if (this.cp.focusMark) {
			this.cp.focusMark.relocate(row,col);
			if(show) this.cp.focusMark.showMark();
		}
	},
	/**
	 * Hides cell's focus mark 
	 */
	hideCellFocus: function () {
		this.focusMark.hideMark();
		if (this.tp.focusMark)
			this.tp.focusMark.hideMark();
		
		if (this.lp.focusMark)
			this.lp.focusMark.hideMark();

		if (this.cp.focusMark)
			this.cp.focusMark.hideMark();
	},
	_realRight: function(top, bottom, right) {
		var minl = right;
		for (var r = bottom; r >= top; --r) {
			var cellL = this.getCell(r, minl);
			if (cellL && cellL.merl < minl) minl = cellL.merl;
		}
		return minl;
	},
	_realLeft: function(top, bottom, left) {
		var maxr = left;
		for (var r = bottom; r >= top; --r) {
			var cellR = this.getCell(r, maxr);
			if (cellR && cellR.merr > maxr) maxr = cellR.merr;
		}
		return maxr;
	},
	/**
	 * Move selection position
	 * @param int key  
	 */
	shiftSelection: function (key) {
		var ls = this.getLastSelection(),
			pos = this.getLastFocus(),
			row = pos.row,
			col = pos.column,
			left = ls.left,
			top = ls.top,
			right = ls.right,
			bottom = ls.bottom,
			update = false,
			seltype = this.selType ? this.selType : zss.SelDrag.SELCELLS;
		
		switch (key) {
		case 'up':
			if (row < bottom)
				bottom--;
			else
				top--;
			break;
		case 'down':
			if (row > top)
				top++;
			else
				bottom++;
			break;
		case 'left':
			right = this._realRight(top, bottom, right);
			if (col < right) {
				right--;
			} else
				left--;
			break;
		case 'right':
			left = this._realLeft(top, bottom, left);
			if (col > left) {
				left++;
			} else
				right++;
			break;
		case 'home':
			right = col;
			left = 0;
			if (seltype == zss.SelDrag.SELALL)
				seltype = zss.SelDrag.SELCOL;
			else if (seltype == zss.SelDrag.SELROW)
				seltype = zss.SelDrag.SELCELLS;
			break;
		case 'end':
			left = col;
			right = this.maxCols - 1;
			if (left == 0) {
				if (seltype == zss.SelDrag.SELCOL)
					seltype = zss.SelDrag.SELALL;
				else if (seltype == zss.SelDrag.SELCELLS)
					seltype = zss.SelDrag.SELROW;
			}
			break;
		case 'pgup':
			if (row < bottom) {
				bottom = bottom - this.pageKeySize;
				if (bottom < row) {
					top = bottom;
					bottom = row;
				}
			} else {
				top -= this.pageKeySize;
			}
			break;
		case 'pgdn':
			if (row > top) {
				top = top + this.pageKeySize;
				if (top > row) {
					bottom = top;
					top = row;
				}
			} else
				bottom += this.pageKeySize; 
			break;
			
		default:
			this.selType = seltype;
			return;
		}
		
		if (left < 0) left = 0;
		if (right >= this.maxCols) right = this.maxCols-1;
		if (top < 0) top = 0;
		if (bottom >= this.maxRows) bottom = this.maxRows-1;
		
		
		//TODO , check cell merge
		//TODO , auto scroll
		
		if (left != ls.left || top != ls.top || right != ls.right || bottom != ls.bottom){
			this.moveCellSelection(left, top, right, bottom, true);
			var ls = this.getLastSelection();
			this.selType = seltype;
			this._sendOnCellSelection(seltype, ls.left, ls.top, ls.right, ls.bottom);
		}
	},
	/**
	 * Sets the highlight area
	 * @param int left
	 * @param int top
	 * @param int right
	 * @param int bottom
	 */
	moveHighlight: function (left, top, right, bottom) {
		//1995691 Highlight doesn't showup after invalidate
		var show = this.hlArea.show;
		if (left < 0 || top < 0 || right < 0 || bottom < 0) {
			this.hideHighlight();
			return;
		}
		var hlRange = new zss.Range(left, top, right, bottom);
		this.hlArea.relocate(hlRange);
		
		if (show)
			this.hlArea.showArea();

		if (this.tp.hlArea) {
			this.tp.hlArea.relocate(hlRange);
			if (show) this.tp.hlArea.showArea();
		}
		if (this.lp.hlArea) {
			this.lp.hlArea.relocate(hlRange);
			if (show) this.lp.hlArea.showArea();
		}
		if (this.cp.hlArea) {
			this.cp.hlArea.relocate(hlRange);
			if (show) this.cp.hlArea.showArea();
		}
	},
	/**
	 * Hides the highlight area
	 */
	hideHighlight: function(){
		this.hlArea.hideArea();
		if (this.tp.hlArea)
			this.tp.hlArea.hideArea();

		if (this.lp.hlArea)
			this.lp.hlArea.hideArea();

		if (this.cp.hlArea)
			this.cp.hlArea.hideArea();
	},
	/**
	 * Animate highlight area
	 * @param boolean start
	 */
	animateHighlight: function (start) {
		this.hlArea.doAnimation(start);
		if (this.tp.hlArea) {
			this.tp.hlArea.doAnimation(start);
		}

		if (this.lp.hlArea) {
			this.lp.hlArea.doAnimation(start);
		}

		if (this.cp.hlArea) {
			this.cp.hlArea.doAnimation(start);
		}
	},
	/**
	 * Display info
	 */
	showInfo: function (text, autohide) {
		this.info.setInfoText(text);
		this.info.showInfo(autohide);
	},
	/**
	 * Hides info
	 */
	hideInfo: function () {
		this.info.hideInfo();
	},
	/**
	 * Display busy
	 */
	showBusy: function (show) {
		jq(this.busycmp).css('visibility', show ? 'visible' : 'hidden');
	},
	/**
	 * Display mark
	 */
	showMask: function (show) {
		jq(this.maskcmp).css('visibility', show ? 'visible' : 'hidden');
	},
	/**
	 * Returns focused cell
	 * @return zss.Cell
	 */
	getFocusedCell: function () {
		var pos = this.getLastFocus();
		return this.getCell(pos.row, pos.column);
	},
	/**
	 * Returns the cell
	 * @param int row row index
	 * @param int col column index
	 * @return zss.Cell
	 */
	getCell: function (row, col) {
		var fzr = this.frozenRow,
			fzc = this.frozenCol,
			cell = null;
		
		if (row <= fzr && col <= fzc)
			cell = this.cp.block.getCell(row, col);
		else if(row <= fzr) 
			cell = this.tp.block.getCell(row, col);
		else if(col <= fzc) 
			cell = this.lp.block.getCell(row, col);
		else 
			cell = this.activeBlock.getCell(row, col);
		return cell;
	},
	/**
	 * Sets block sync event to server
	 */
	sendSyncblock: function (now) {
		var spcmp = this.sp.comp,
			dp = this.dp,
			brange = this.activeBlock.range;

		this._wgt.fire('onZSSSyncBlock', {
			sheetId: this.sheetid,
			dpWidth: dp.width,
			dpHeight: dp.height,
			viewWidth: spcmp.clientWidth,
			viewHeight: spcmp.clientHeight,
			blockLeft: brange.left,
			blockTop: brange.top,
			blockRight: brange.right,
			blockBottom: brange.bottom,
			fetchLeft: -1,
			fetchTop: -1,
			fetchWidth: -1,
			fetchHeight: -1,
			rangeLeft: -1,
			rangeTop: -1,
			rangeRight: -1,
			rangeBottom: -1
		}, now ? {toServer: true} : null, (now ? 25 : -1));
	},
	_insertNewColumn: function (col, size, extnm) {
		this.activeBlock.insertNewColumn(col,size);
		var fzc = this.frozenCol;
		if(col <= fzc){
			this.lp.insertNewColumn(col, size);
			this.cp.insertNewColumn(col, size, extnm);
		}
		this.tp.insertNewColumn(col, size, extnm);	
	},
	_insertNewRow: function (row, size, extnm) {
		this.activeBlock.insertNewRow(row, size);
		var fzr = this.frozenRow;
		if (row <= fzr) {
			this.tp.insertNewRow(row, size);
			this.cp.insertNewRow(row, size, extnm);
		}
		this.lp.insertNewRow(row, size, extnm);
			
	},
	_removeColumn: function (col, size, extnm) {
		this.activeBlock.removeColumn(col, size);
		var fzc = this.frozenCol;
		if (col <= fzc) {
			this.lp.removeColumn(col, size);
			this.cp.removeColumn(col, size, extnm);
		}
		this.tp.removeColumn(col, size, extnm);
	},
	_removeRow: function (row, size, extnm) {
		this.activeBlock.removeRow(row, size);
		var fzr = this.frozenRow;
		if (row <= fzr) {
			this.tp.removeRow(row, size);
			this.cp.removeRow(row, size, extnm);
		}
		this.lp.removeRow(row, size, extnm);
	},
	_removeMergeRange: function (result) {
		var id = result.id,
			left = result.left,
			top = result.top,
			right = result.right,
			bottom = result.bottom,
			sheetid = this.sheetid;
		
		var cssid = sheetid + "-sheet" + ((zk.opera) ? "-opera" : ""),//opera bug, it cannot insert rul to special position
			name = "#" + sheetid;
		zcss.removeRule(name + " .zsmerge" + id, cssid);
		zcss.removeRule(name + " .zsmerge" + id + " .zscelltxt", cssid);
		
		this.mergeMatrix.removeMergeRange(id);
		this.activeBlock.removeMergeRange(id, left, top, right, bottom);

		if(this.cp.block)
			this.cp.block.removeMergeRange(id, left, top, right, bottom);
		if(this.tp.block)
			this.tp.block.removeMergeRange(id, left, top, right, bottom);
		if(this.lp.block)
			this.lp.block.removeMergeRange(id, left, top, right, bottom);
	},
	_addMergeRange: function (result) {
		var id = result.id,
			left = result.left,
			top = result.top,
			right = result.right,
			bottom = result.bottom,
			width = result.width;

		var cssid = this.sheetid + "-sheet" + ((zk.opera) ? "-opera" : ""),//opera bug, it cannot insert rul to special position
			cp = this.cellPad,
			celltextwidth = width - 2 * cp - 1,
			cellwidth = zk.ie || zk.safari || zk.opera ? celltextwidth : width,
			name = "#" + this.sheetid;
		zcss.setRule(name + " .zsmerge" + id, "width", cellwidth + "px", true, cssid);
		zcss.setRule(name + " .zsmerge" + id + " .zscelltxt", "width", celltextwidth + "px", true, cssid);
		
		this.mergeMatrix.addMergeRange(id, left, top, right, bottom);		
		this.activeBlock.addMergeRange(id, left, top, right, bottom);

		if(this.cp.block)
			this.cp.block.addMergeRange(id, left, top, right, bottom);
		if(this.tp.block)
			this.tp.block.addMergeRange(id, left, top, right, bottom);
		if(this.lp.block)
			this.lp.block.addMergeRange(id, left, top, right, bottom);
	}
}, {
	NOFOCUS: 0,
	FOCUSED: 2,
	START_EDIT: 5, //2*2 + 1; //async state is odd
	EDITING: 6, //3*2 ,
	STOP_EDIT: 9, //4*2 + 1;//async state is odd
	_initInnerComp: function (sheet) {
		var wgt = sheet._wgt;
		sheet.maskcmp = wgt.$n('mask');
		sheet.busycmp = wgt.$n('busy');
		sheet.spcmp = wgt.$n('sp');//scroll panel comp
		sheet.topcmp = wgt.$n('top');//top panel comp
		sheet.leftcmp = wgt.$n('left');//left panel comp
		sheet.dpcmp = wgt.$n('dp');//data panel comp
		sheet.wpcmp = wgt.$n('wp');//widget panel comp
		sheet.sinfocmp = wgt.$n('sinfo');
		sheet.infocmp = wgt.$n('info');
		sheet.cpcmp = wgt.$n('co');


		sheet.sp = new zss.ScrollPanel(sheet);
		sheet.dp = new zss.DataPanel(sheet);
		sheet.tp = new zss.TopPanel(sheet, sheet.topcmp);
		sheet.lp = new zss.LeftPanel(sheet, sheet.leftcmp);
		sheet.cp = new zss.CornerPanel(sheet);
		
		var dppadcmp = wgt.$n('datapad'),
			blockcmp = wgt.$n('block');
	
		sheet.activeBlock = new zss.MainBlockCtrl(sheet, blockcmp, 0, 0);
		sheet.activeBlock.loadByComp(blockcmp);
		
		var next = wgt.$n('select');
		if (jq(next).attr('zs.t') == "SSelect") {
			sheet.selareacmp = next;
			sheet.selchgcmp = wgt.$n('selchg');
			sheet.focusmarkcmp = wgt.$n('focmark');
			sheet.hlcmp = wgt.$n('highlight');
			sheet.editorcmp = wgt.$n('eb');
			
			sheet.focusMark = new zss.FocusMarkCtrl(sheet, sheet.focusmarkcmp, sheet.initparm.focus.clone());
			sheet.selArea = new zss.SelAreaCtrl(sheet, sheet.selareacmp, sheet.initparm.selrange.clone());
			sheet.selChgArea = new zss.SelChgCtrl(sheet, sheet.selchgcmp);
			sheet.hlArea = new zss.Highlight(sheet, sheet.hlcmp,sheet.initparm.hlrange.clone(), "inner");
			sheet.editor = new zss.Editbox(sheet);
			
		} else {
			//error
			//zk.log('error to parse component');
		}
		
		
		//initial scroll info
		sheet.sinfo = new zss.ScrollInfo(sheet, sheet.sinfocmp);
		sheet.info = new zss.Info(sheet, sheet.infocmp);
	},
	_getColumnCells: function (sheet, index) {
		var cells = [],
			next = jq(sheet.dpcmp).children("DIV:first")[0];
		while (next) {
			if (jq(next).attr('zs.t') == "SRow") {
				var nextcell = jq(next).children("DIV:first")[0];
				while (nextcell) {
					if (nextcell.ctrl.c == index) {
						cells.push(nextcell);
						break;
					}
					nextcell = jq(nextcell).next("DIV")[0];
				}
			}
			next = jq(next).next("DIV")[0];	
		}
		return cells;
	},
	_getVisibleRange: function (sheet) {
		var sp = sheet.sp,
			spcmp = sp.comp,
			block = sheet.activeBlock,
			leftCell = block.range.left,
			bottomCell = block.range.bottom,
			scrollLeft = spcmp.scrollLeft,
			scrollTop = spcmp.scrollTop,
			custColWidth = sheet.custColWidth,
			custRowHeight = sheet.custRowHeight,
			viewWidth = spcmp.clientWidth -  sheet.leftWidth,
			viewHeight = spcmp.clientHeight - sheet.topHeight,	
			left = custColWidth.getCellIndex(scrollLeft)[0],
			top = custRowHeight.getCellIndex(scrollTop)[0],
			right = custColWidth.getCellIndex(scrollLeft + viewWidth)[0],
			bottom = custRowHeight.getCellIndex(scrollTop + viewHeight)[0];

		if (right > sheet.maxCols - 1) right = sheet.maxCols - 1;
		if (bottom > sheet.maxRows - 1) bottom = sheet.maxRows - 1; 

		return new zss.Range(left, top, right, bottom);
	},
	/**
	 * get cell position(row,col) according to given client position (x,y) of browser and current UI display
	 * @param zss.SSheetCtrl
	 * @param int x page offset
	 * @param int y page offset
	 * @param boolean ignorefrezon don't care if over a frezon panel 
	 */
	_calCellPos: function (sheet, x, y, ignorefrezon) {
		var row = col = -1,
			dpofs = zk(sheet.dp.comp).revisedOffset(),
			custColWidth = sheet.custColWidth,
			custRowHeight = sheet.custRowHeight,
			rx = x,
			ry = y,
			fzr = sheet.frozenRow,
			fzc = sheet.frozenCol;
		
		if (!ignorefrezon && (fzr > -1 || fzc > -1)) {
			var sheetofs = zk(sheet.comp).revisedOffset(),
				fx = fy = -1;
			if (fzc > -1)
				fx = custColWidth.getStartPixel(fzc + 1);
			if (fzr > -1)
				fy = custRowHeight.getStartPixel(fzr + 1);

			rx = x - sheetofs[0] - sheet.leftWidth;
			ry = y - sheetofs[1] - sheet.topHeight;
			if (rx > fx && ry > fy) {
				rx = x - dpofs[0] - sheet.leftWidth;
				ry = y - dpofs[1] - sheet.topHeight;
			} else if (ry > fy)
				ry = y - dpofs[1] - sheet.topHeight;
			else if(rx > fx)
				rx = x - dpofs[0] - sheet.leftWidth;

		} else {
			rx = x - dpofs[0] - sheet.leftWidth;
			ry = y - dpofs[1] - sheet.topHeight;
		}
		
		var xcol = custColWidth.getCellIndex(rx),
			xrow = custRowHeight.getCellIndex(ry);
		return [xrow[0], xcol[0], rx, ry, xcol[1], xrow[1]];
	},
	/* static, get current sheet obj */
	_curr: function (obj) {
		var cmp = zss.SSheetCtrl._currcmp(obj);
		if (cmp != null && cmp.ctrl) return cmp.ctrl;
		return null;
	},
	_currcmp: function(obj) {
		if (typeof obj == "string")
			return zk.Widget.$(obj).$n();

		if (obj.sheetid)
			return zk.Widget.$(obj.sheetid).$n();

		if (obj.ctrl)
			return zk.Widget.$(obj.ctrl.sheetid).$n();

		return null;
	},
	_assignSheet: function (ctrl) {
		ctrl.sheet = zss.SSheetCtrl._curr(ctrl.sheetid);
	}
});
})();