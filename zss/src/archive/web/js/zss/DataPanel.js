/* DataPanel.js

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
 * A DataPanel represent the spreadsheet cells 
 */
zss.DataPanel = zk.$extends(zk.Object, {
	cacheMap: {}, //handle element cache, not implement yet
	$init: function (sheet) {
		this.$supers('$init', arguments);
		var wgt = sheet._wgt,
			dataPanel = wgt.$n('dp'),
			focus  = this.focustag = wgt.$n('fo');

		this.id = dataPanel.id;
		this.sheetid = sheet.sheetid;
		this.sheet = sheet;
		this.comp = dataPanel;
		this.comp.ctrl = this;
		this.padcomp =  wgt.$n('datapad');

		wgt.domListen_(focus, 'onBlur', '_doDataPanelBlur');
		wgt.domListen_(focus, 'onFocus', '_doDataPanelFocus');

		this.width = zk.parseInt(jq(dataPanel).attr("z.w"));
		this.height = zk.parseInt(jq(dataPanel).attr("z.h"));

		this.paddingl = this.paddingt = 0;
		var pdl = this.paddingl + sheet.leftWidth,
			pdt = sheet.topHeight;//this.paddingt + sheet.topHeight;
		jq(dataPanel).css({'padding-left': jq.px0(pdl), 'padding-top': jq.px0(pdt), 'width': jq.px0(this.width), 'height': jq.px0(this.height)});
	},
	cleanup: function() {
		var focus = this.focustag,
			wgt = this.sheet._wgt;
		wgt.domUnlisten_(focus, 'onBlur', '_doDataPanelBlur');
		wgt.domUnlisten_(focus, 'onFocus', '_doDataPanelFocus');
		
		this.invalid = true;
		this.cacheMap = this.focustag = null;

		if (this.comp) this.comp.ctrl = null;
		this.comp = this.padcomp = this.sheet = null;
	},
	/**
	 * Sets the height
	 * @param int diff
	 */
	updateHeight: function (diff) {
		this.height += diff;
		jq(this.comp).css('height', jq.px0(this.height));
	},
	/**
	 * Sets the width
	 * @param int diff
	 */
	updateWidth: function (diff) {
		this.width += diff;
		jq(this.comp).css('width', jq.px0(this.width));
	},
	_fixSize: function (block) {
		var sheet = this.sheet,
			custColWidth = sheet.custColWidth,
			custRowHeight = sheet.custRowHeight;
		
		this.paddingl = custColWidth.getStartPixel(block.range.left);
		this.paddingt = custRowHeight.getStartPixel(block.range.top);
		
		var width = this.width - this.paddingl,
			height = this.height,
			pdl = this.paddingl + sheet.leftWidth;
		
		jq(this.comp).css({'padding-left': jq.px0(pdl), 'padding-top': jq.px0(sheet.topHeight), 'width': jq.px0(width)});
		jq(this.padcomp).css('height', jq.px0(this.paddingt));
		sheet.tp._updateLeftPadding(pdl);
		sheet.lp._updateTopPadding(this.paddingt);
	},
	/**
	 * Returns whether is focus on frozen area or not.
	 * @return boolean
	 */
	isFocusOnFrozen: function () {
		var sheet = this.sheet,
			fzr = sheet.frozenRow,
			fzc = sheet.frozenCol,
			pos = sheet.getLastFocus();

		return fzc > -1 ? (pos.column <= fzc || pos.row <= fzr) ? true : false : false;
	},
	/**
	 * Start editing on cell
	 */
	startEditing: function (evt, val) {
		var sheet = this.sheet;
		if (sheet.config.readonly) 
			return false;

		var pos = sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		
		if (this.isFocusOnFrozen()) {
			sheet.showInfo("Can not edit on a frozen cell.", true);
			return false;	
		}
		if (!val) val = null;
		
		//bug, if I use loadCell and didn't add a call back function to scroll to loaded block
		//when after loadCell, it always invoke loadForVisible , then UI will invoke a loadForVisible 
		//depends on current UI range. so UI moveFocus instead.
		//sheet.activeBlock.loadCell(row,col,5,null);
		sheet.dp.moveFocus(row, col, true, true);
		
		sheet.state = zss.SSheetCtrl.START_EDIT;
		sheet._wgt.fire('onStartEditing',
				{token: "", sheetId: sheet.serverSheetId, row: row, col: col, clienttxt: val}, null, 25);
		return true;
	},
	//bug #117 Barcode Scanner data incomplete
	_insertAndNext: function (row, col, value) {
		var sheet = this.sheet;
		sheet.state = zss.SSheetCtrl.FOCUSED;
		sheet.dp.moveDown();
		sheet._wgt.fire('onStopEditing', 
				{token: null, sheetId: sheet.serverSheetId, row: row, col: col, value: value}, {toServer: true}, 25);
	},
	/* move focus to a cell and start editing the value. This method will do - 
	* 1. check if the cell loaded already or not
	* 2. if not loaded, it will cause a asynchronized loading. after loading then do 3.
	* 3. if loaded, then invoke _moveFocus to move to loaded cell; then call sheet.editor.edit() to start editing
	*/
	//bug #117 Barcode Scanner data incomplete
	_moveFocusAndEdit: function(row, col, value, focus) {
		var sheet = this.sheet,
			fzr = sheet.frozenRow,
			fzc = sheet.frozenCol,
			local = this,
			fn = function () {
				if (focus)
					local._moveFocus(row, col, true, true);
				if (sheet.state != zss.SSheetCtrl.START_EDIT)
					local.startEditing();
				cell = sheet.getFocusedCell();
				if (cell != null) {
					sheet.editor.edit(cell.comp, row, col, value);
					sheet.state = zss.SSheetCtrl.EDITING;
				}
			},
			block = sheet.activeBlock,
		//if target cell is on frozon row and col
		//then modify the load target to a loaded cell
			lr = (row <= fzr ? block.range.top : row),
			lc = (col <= fzc ? block.range.left : col),
			r = block.loadCell(lr, lc, 5, fn);
		if(r)
			fn();
	},
	//@param value to start editing
	//@param server boolean whether the value come from server 
	_startEditing: function (value, server) {
		var sheet = this.sheet,
			cell = sheet.getFocusedCell();

		//bug #117 Barcode Scanner data incomplete
		if (cell != null) {
			var editor = sheet.editor;
			if( sheet.state == zss.SSheetCtrl.START_EDIT) {
				if (!server)
					value = sheet._clienttxt;
				sheet._clienttxt = '';
				var pos = sheet.getLastFocus(),
					row = pos.row,
					col = pos.column,
					j = value.indexOf(zkS._enterChar),
					val;
				//handle speed input with Enter(0x0D) key
				while (j >= 0) {
					val = value.substring(0, j);
					if (sheet.state != zss.SSheetCtrl.START_EDIT) {
						sheet.state = zss.SSheetCtrl.START_EDIT;
						sheet._wgt.fire('onStartEditing',
								{token: "", sheetId: sheet.serverSheetId, row: row, col: col, clienttxt: val}, null, 25);
					}
					sheet.state = zss.SSheetCtrl.FOCUSED;
					sheet._wgt.fire('onStopEditing', 
							{token: "", sheetId: sheet.serverSheetId, row: row, col: col, value: val}, {toServer: true}, 25);
					++row; //move down to next cell
					value = value.substring(j+1);
					j = value.indexOf(zkS._enterChar);
				}
				if (value.length > 0)
					this._moveFocusAndEdit(row, col, value, pos.row != row); //set focus and enter editing mode
				else if (pos.row != row)
					this.moveFocus(row, col, true, true);
			} else if (server && sheet.state == zss.SSheetCtrl.EDITING && editor.comp.value != value) {
				var pos = sheet.getLastFocus();
				editor.edit(cell.comp, pos.row, pos.column, value);
			}
		}
	},
	//optimize F2
	_openEditbox: function () {
		var sheet = this.sheet,
			cell = sheet.getFocusedCell();
			
		if (cell != null) {
			if( sheet.state == zss.SSheetCtrl.START_EDIT) {
				var editor = sheet.editor,
					pos = sheet.getLastFocus(),
					value = cell.edit;
				editor.edit(cell.comp, pos.row, pos.column, value ? value : '');
				sheet.state = zss.SSheetCtrl.EDITING;
			}
		}
	},
		
	/**
	 * Cancel editing on cell
	 */
	cancelEditing: function () {
		var sheet = this.sheet;

		if (sheet.state >= zss.SSheetCtrl.FOCUSED) {
			var editor = sheet.editor;
			editor.cancel();
			sheet.state = zss.SSheetCtrl.FOCUSED;
			this.reFocus(true);
		}
	},
	/**
	 * Stop editing on cell
	 */
	stopEditing: function (type) {
		var sheet = this.sheet;
		if (sheet.state == zss.SSheetCtrl.EDITING) {

			sheet.state = zss.SSheetCtrl.STOP_EDIT;

			var editor = this.sheet.editor;
			editor.disable(true);
			var row = editor.row,
				col = editor.col,
				value = editor.currentValue(),
				cell = sheet.getCell(row, col),
				edit = cell ? cell.edit : null;
			
			if (edit != value) { //has to send back to server
				var token = "";
				if (type == "movedown") {//move down after aysnchronized call back
					token = zkS.addCallback(function(){
						if (sheet.invalid) return;
						sheet.state = zss.SSheetCtrl.FOCUSED;
						sheet.dp.moveDown();
					});
				} else if (type == "moveright") {//move right after aysnchronized call back
					token = zkS.addCallback(function(){
						if (sheet.invalid) return;
						sheet.state = zss.SSheetCtrl.FOCUSED;
						sheet.dp.moveRight();
					});
				} else if (type == "moveleft") {//move right after aysnchronized call back
					token = zkS.addCallback (function(){
						if (sheet.invalid) return;
						sheet.state = zss.SSheetCtrl.FOCUSED;
						sheet.dp.moveLeft();
					});
				} else if (type == "refocus") {//refocuse after aysnchronized call back
					token = zkS.addCallback(function(){
						if (sheet.invalid) return;
						sheet.state = zss.SSheetCtrl.FOCUSED;
						sheet.dp.reFocus(true);
					});
				} else if (type == "lostfocus") {//lost focus after aysnchronized call back
					token = zkS.addCallback(function(){
						if (sheet.invalid) return;
						sheet.dp._doFocusLost();
					});
				}
	
				//zkau.send({uuid: this.sheetid, cmd: "onZSSStopEditing", data: [token,sheet.serverSheetId,row,col,value]},25/*zkau.asapTimeout(cmp, "onOpen")*/);
				sheet._wgt.fire('onStopEditing', 
						{token: token, sheetId: sheet.serverSheetId, row: row, col: col, value: value}, {toServer: true}, 25);
			} else {
				this._stopEditing();
				if (type == "movedown") {//move down
					if (sheet.invalid) return;
					sheet.state = zss.SSheetCtrl.FOCUSED;
					this.moveDown();
				} else if (type == "moveright") {//move right
					if (sheet.invalid) return;
					sheet.state = zss.SSheetCtrl.FOCUSED;
					this.moveRight();
				} else if (type == "moveleft") {//move right
					if (sheet.invalid) return;
					sheet.state = zss.SSheetCtrl.FOCUSED;
					this.moveLeft();
				} else if (type == "refocus") {//refocuse
					if (sheet.invalid) return;
					sheet.state = zss.SSheetCtrl.FOCUSED;
					this.reFocus(true);
				} else if (type == "lostfocus") {//lost focus
					if (sheet.invalid) return;
					this._doFocusLost();
				}
			}
		}
	},
	_stopEditing: function () {
		var sheet = this.sheet;
		if (sheet.state == zss.SSheetCtrl.STOP_EDIT)
			var editor = this.sheet.editor.stop();
	},
	/* move focus to a cell, this method will do - 
	* 1. stop editing, 
	* 2. check if the cell loaded already or not
	* 3. if not loaded, it will cause a asynchronized loading. after loading then do 4.
	* 4. if loaded, then invoke _moveFocus to move to loaded cell
	* 
	* @param {int} row
	* @param {int} col
	* @param {bollean} scroll scroll the cell in view after loading
	* @param {boolean} selection move selection to focus cell, (the selection will change)
	* @param {boolean} noevt don't send onCellFocused Event to server side
	* @param {boolean} noslevt don't send onCellSelection Event to server side
	*/
	moveFocus: function(row, col, scroll, selection, noevt, noslevt) {
		
		this.stopEditing("refocus");
		var sheet = this.sheet,
			fzr = sheet.frozenRow,
			fzc = sheet.frozenCol,
			local = this,
			fn = function () {
				//TODO, cell not initial, what sholud i do?
				local._moveFocus(row, col, scroll, selection, noevt, noslevt);
			},
			block = sheet.activeBlock,
		//if target cell is on frozon row and col
		//then modify the load target to a loaded cell
			lr = (row <= fzr ? block.range.top : row),
			lc = (col <= fzc ? block.range.left : col),
			r = block.loadCell(lr, lc, 5, fn);
		if(r)
			fn();
	},
	/* move focus to new cell , it will check cell is initialed or not(delay init)*/
	_moveFocus: function (row, col, scroll, selection, noevt, noslevt) {
		var sheet = this.sheet,
			cell = sheet.getCell(row, col);
		
		if (!cell)
			return false;	
		var cellcmp = cell.comp,
			ml = cell.merl;//this cell is merged by a left cell. cellcmp.ctrl.merl;
		if (zkS.t(ml) && ml != col)
			return this._moveFocus(row, ml, scroll, selection, noevt, noslevt);
				
		sheet.moveCellFocus(row, col);
		if (!noevt) sheet._sendOnCellFocused(row, col); 
		if (selection) {
			sheet.moveCellSelection(col, row, col, row);
			var ls = sheet.getLastSelection(); 
			if (!noslevt) sheet._sendOnCellSelection(zss.SelDrag.SELCELLS, ls.left, ls.top, ls.right, ls.bottom);
		}
		
		var sheet = this.sheet;
		this._gainFocus(true, noevt, sheet.state < zss.SSheetCtrl.FOCUSED ? false : true);
		
		if (scroll)
			sheet.sp.scrollToVisible(null, null, cell);	
		if (sheet.config.textOverflow)
			zss.Cell._processOverflow(cell);
		return true;
	},
	/* re-focus current cell , and stop editing*/
	reFocus: function (scroll) {
		
		var pos = this.sheet.getLastFocus();
		this.moveFocus(pos.row, pos.column, scroll, true);
	},
	/** gain focus to the focus tag , and stop editing
	 *  @param trigger should fire focustag.focus (if genFocus is call by a focustag listener for focus, then it is false 
	 */
	gainFocus: function (trigger) {
		this.stopEditing("refocus");
		this._gainFocus(trigger);
	},
	/* gain focus to the focus cell*/
	_gainFocus: function (trigger, noevt, noslloc) {
		var sheet = this.sheet,
			local = this,
			pos = sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;

		if (sheet.state == zss.SSheetCtrl.NOFOCUS && !noevt)
			sheet._sendOnCellFocused(row, col); 
		
		if (sheet.state < zss.SSheetCtrl.FOCUSED)
			sheet.state = zss.SSheetCtrl.FOCUSED;
		
		if (!noslloc) {
			sheet.moveCellFocus(row, col);
			var ls = sheet.selArea.lastRange;
			sheet.moveCellSelection(ls.left, ls.top, ls.right, ls.bottom);
		}
		
		var lhl = sheet.hlArea.lastRange;
		sheet.moveHighlight(lhl.left, lhl.top, lhl.right, lhl.bottom);
		sheet.animateHighlight(true);

		if (trigger && sheet.state == zss.SSheetCtrl.FOCUSED){ 
			setTimeout(function () {
				local.focustag.focus();
			}, 0);
		}
	},
	/* process focus lost */
	_doFocusLost: function() {
		var sheet = this.sheet;
		sheet.state = zss.SSheetCtrl.NOFOCUS;
		sheet.hideCellFocus();
		sheet.hideCellSelection();
		sheet.animateHighlight(false);
		sheet._wgt.fire('onBlur');
	},
	/**
	 * Select cell
	 * @param int row the row
	 * @param int col the column
	 */
	selectCell: function(row, col, forcemove) {
		var sheet = this.sheet;
		if(sheet.maxRows <= row || sheet.maxCols <= col) return;

		if ((forcemove || sheet.state >= zss.SSheetCtrl.FOCUSED) && row >-1 && col > -1) {
			//sheet has focus move, load cell to client,and scroll the such cell
			this.moveFocus(row, col, true, true);
		} else {//user click back to the sheet, don't move the scroll
			//gain focus and reallocate mark , then show it, 
			//don't move the focus because the cell maybe doens't exist in block anymore.
			sheet.moveCellFocus(row, col);
			sheet.moveCellSelection(col, row, col, row);			
			this._gainFocus(true, false, true);
		}
	},
	/**
	 * Move page up
	 */
	movePageup: function (evt) {
		if (zkS.isEvtKey(evt, "s")) {
			this.sheet.shiftSelection("pgup");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if (row < 0 || col < 0) return;
		var sheet = this.sheet;
		if (row > 0) {
			//TODO: calculate the first row in previous page
			var prevrow = row - sheet.pageKeySize;
			row = prevrow;
			row = row < 0 ? 0 : row;
		}
		this.moveFocus(row, col, true, true);
	},
	/**
	 * Move page down
	 */
	movePagedown: function (evt) {
		if (zkS.isEvtKey(evt,"s")){
			this.sheet.shiftSelection("pgdn");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if(row <0 || col < 0) return;
		var sheet = this.sheet;
		if (row < sheet.maxRows - 1) {
			//TODO: calculate the first row of next page
			var nextrow = row + sheet.pageKeySize;
			row = nextrow;
			if (row > sheet.maxRows - 1) {
				row = sheet.maxRows - 1;
			}
		}
		this.moveFocus(row, col, true, true);
	},
	/**
	 * Move to the end column
	 */
	moveEnd: function (evt) {
		if (zkS.isEvtKey(evt, "s")) {
			this.sheet.shiftSelection("end");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if (row < 0 || col < 0) return;

		var sheet = this.sheet;
		if (col < sheet.maxCols -1)
			col = sheet.maxCols -1;

		this.moveFocus(row, col, true, true);
	},
	/**
	 * Move to the first column 
	 */
	moveHome: function (evt) {
		if (zkS.isEvtKey(evt, "s")) {
			this.sheet.shiftSelection("home");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if (row < 0 || col < 0) return;
		if (col > 0)
			col = 0;

		this.moveFocus(row, col, true, true);
	},
	/**
	 * Move up
	 */
	moveUp: function (evt) {
		if (zkS.isEvtKey(evt,"s")) {
			this.sheet.shiftSelection("up");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if (row < 0 || col < 0) return;
		if (row > 0)
			--row;
		this.moveFocus(row, col, true, true);
	},
	/**
	 * Move down
	 */
	moveDown: function(evt) {
		if (zkS.isEvtKey(evt, "s")) {
			this.sheet.shiftSelection("down");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if (row <0 || col < 0) return;
		var sheet = this.sheet;
		if (row < sheet.maxRows - 1)
			++row;

		this.moveFocus(row, col, true, true);
	},
	/**
	 * Move to the left column
	 */
	moveLeft: function(evt) {
		if (zkS.isEvtKey(evt,"s")) {
			this.sheet.shiftSelection("left");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if (row < 0 || col < 0) return;
		if (col > 0)
			--col;
		this.moveFocus(row, col, true, true);
	},
	/**
	 * Move to the right column 
	 */
	moveRight: function(evt) {
		if (zkS.isEvtKey(evt, "s")) {
			this.sheet.shiftSelection("right");
			return;
		}
		var pos = this.sheet.getLastFocus(),
			row = pos.row,
			col = pos.column;
		if(row < 0 || col < 0) return;
		
		var sheet = this.sheet,
			cell = sheet.getCell(row, col);
		if (cell) {
			var mr = cell.merr;
			if(zkS.t(mr))
				col = mr;
		}
		if (col < sheet.maxCols - 1)
			col++;
		
		this.moveFocus(row, col, true, true);
	}
});