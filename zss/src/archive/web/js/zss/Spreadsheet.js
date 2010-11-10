/* Spreadsheet.js

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
	function _calScrollWidth (el) {
		if(zkS.t(zss.Spreadsheet.scrollWidth)) return;
	    // scroll scrolling div
	    var scr = document.createElement('div');
	    scr.style.position = 'absolute';
	    scr.style.top = '0px';
	    scr.style.left = '0px';
	    scr.style.width = '50px';
	    scr.style.height = '50px';
	    scr.style.overflow = 'auto';

	    // scrolli content div
	    var inn = document.createElement('div');
	    inn.style.width = '100px';
	    inn.style.height = '100px';

	    // Put the scrolli div in the scrolling div
	    scr.appendChild(inn);

	    // Append the scrolling div to the doc
	    el.appendChild(scr);

		var ow = scr.offsetWidth;
		var cw = scr.clientWidth;

	    //calcuate the scrollWidth;	
		var ow = scr.offsetWidth,
			cw = scr.clientWidth,
			oh = scr.offsetHeight,
			ch = scr.clientHeight;
		el.removeChild(scr);
		zss.Spreadsheet.scrollWidth = (ow - cw);
		if (zss.Spreadsheet.scrollWidth == 0)
			zss.Spreadsheet.scrollWidth = oh - ch;
		return;
	}
	
	function _doCellUpdateCmd (sheet, data, token) {
		sheet._cmdCellUpdate(jq.evalJSON(data));
		if (token != "") {
			zkS.doCallback(token);
		}
	}
	function _doBlockUpdateCmd (sheet, data, token) {
		sheet._cmdBlockUpdate(jq.evalJSON(data));
		if (token != "")
			zkS.doCallback(token);
	}
	function _doInsertRCCmd (sheet, data, token) {
		sheet._cmdInsertRC(jq.evalJSON(data));
		if (token != "")
			zkS.doCallback(token);	
	}
	function _doRemoveRCCmd (sheet, data, token) {
		sheet._cmdRemoveRC(jq.evalJSON(data), true);
		if (token != "")
			zkS.doCallback(token);
	}
	function _doMergeCmd (sheet, data, token) {
		sheet._cmdMerge(jq.evalJSON(data), true);
		if (token != "")
			zkS.doCallback(token);
	}
	function _doSizeCmd (sheet, data, token) {
		sheet._cmdSize(jq.evalJSON(data), true);
		if (token != "")
			zkS.doCallback(token);
	}
	function _doMaxrowCmd (sheet, data) {
		sheet._cmdMaxrow(data);
	}
	function _doMaxcolumnCmd (sheet, data) {
		sheet._cmdMaxcolumn(data);
	}
	
	/**
	 * Create a StyleSheet at the end of head node
	 * @param {String} cssText the default style css text.
	 * @param {String} id the id of this style sheet, null-able, you must assign the id if you want to remove it later.
	 * @return {Object} StyleSheet object
	 */
	function createSSheet (cssText, id) {
	    var head = document.getElementsByTagName("head")[0],
    		sheetele = document.createElement("style"),
    		sheetobj;
	    jq(sheetele).attr("type", "text/css");
		
	    if (id)
	       jq(sheetele).attr("id", id);
	    if (zk.ie) {
	        head.appendChild(sheetele);
	        sheetobj = sheetele.styleSheet;
	        sheetobj.cssText = cssText;
	    } else {
	        try {
	            sheetele.appendChild(document.createTextNode(cssText));
	        } catch (e) {
	            sheetele.cssText = cssText;
	        }
	        head.appendChild(sheetele);
			sheetobj = _getElementSheet(sheetele);
	    }
	    return sheetobj;
	}
	/**
	 * remove StyleSheet object by id if exist.
	 * @param {String} id the style sheet id 
	 */
	function removeSSheet (id) {
		var node = document.getElementById(id);
		if(node && node.type == "text/css"){
			node.parentNode.removeChild(node);
			if (zk.ie) {//refresh for IE
				createSSheet("", id);
				node = document.getElementById(id);
				node.parentNode.removeChild(node);		
			}
		}
	}
	/**
	 * get stylesheet object from a element
	 */
	function _getElementSheet () {
		if (zk.ie) {
			return function (element) {
				return element.styleSheet;
			};
		} else {
			return function (element) {
				return element.sheet;
			};
		}
	}
	
	/**
	 * Synchronize frozen area with top panel, left panel and corner panel
	 */
	function syncFrozenArea (sheet) {
		var scroll = sheet.sp.comp,
			corner = sheet.cp.comp,
			currentWidth =  corner.offsetWidth,
			maxWidth = scroll.clientWidth,
			currentHeight = corner.offsetHeight,
			maxHeight = scroll.clientHeight;
		
		if (currentWidth > maxWidth) {
			var max = jq.px0(maxWidth),
				left = sheet.lp.comp;
			jq(corner).css('width', max);
			jq(left).css('width', max);
			//jq(scroll).css('overflow-x', 'hidden');
		}

		if (currentHeight > maxHeight) {
			var height = jq.px0(maxHeight),
				top = sheet.tp.comp;
			jq(corner).css('height', height);
			jq(top).css('height', height);
			//jq(scroll).css('overflow-y', 'hidden');
		}
	}
/**
 * zkSSheet -> zss.Spreadsheet
 */
zss.Spreadsheet = zk.$extends(zul.Widget, {
	$define: {
		/**
		 * synchronized update data
		 * @param array
		 */
		dataBlockUpdate: _dataUpdate = function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			var token = v[0],
				data = v[2];
			
			if (sheet._initiated)
				_doBlockUpdateCmd(sheet, data, token);
			else
				sheet.addSSInitLater(_doBlockUpdateCmd, sheet, data, token);
		},
		dataBlockUpdateJump: _dataUpdate,
		dataBlockUpdateEast: _dataUpdate,
		dataBlockUpdateWest: _dataUpdate,
		dataBlockUpdateSouth: _dataUpdate,
		dataBlockUpdateNorth: _dataUpdate,
		dataBlockUpdateError: _dataUpdate,
		/**
		 * update data
		 * @param array
		 */
		dataUpdate: _updateCell = function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet)	return;
			var token = v[0],
				data = v[2];
			if (sheet._initiated)
				_doCellUpdateCmd(sheet, data, token);
			else
				sheet.addSSInitLater(_doCellUpdateCmd, sheet, data, token);
		},
		dataUpdateStart: _updateCell,
		dataUpdateCancel: _updateCell,
		dataUpdateStop: _updateCell,
		redrawWidget: function (v) {
			var serverSheetId = v[0],
				wgtUuid = v[1],
				sheet = this.sheetCtrl;
			if (!sheet || serverSheetId != this.getSheetId())
				return;

			var wgt = zk.Widget.$(wgtUuid);
			if (wgt)
				wgt.redrawWidgetTo(sheet);
		},
		/**
		 * rename. insertrc_ -> insertRowColumn
		 */
		insertRowColumn: function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			
			var token = v[0],
				data = v[2],
				sheet = this.sheetCtrl;
			if (sheet._initiated)
				_doInsertRCCmd(sheet, data, token);
			else
				sheet.addSSInitLater(_doInsertRCCmd, sheet, data, token);
		},
		/**
		 * removerc_ -> removeRowColumn
		 */
		removeRowColumn: function (v) {
			var sheet = sheet = this.sheetCtrl;
			if (!sheet) return;
			var token = v[0], 
				data = v[2];

			if (sheet._initiated)
				_doRemoveRCCmd(sheet, data, token);
			else
				sheet.addSSInitLater(_doRemoveRCCmd, sheet, data, token);
		},
		/**
		 * merge_ -> mergeCell
		 */
		mergeCell: function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			var token = v[0], 
				data = v[2];

			if (sheet._initiated)
				_doMergeCmd(sheet, data, token);
			else
				sheet.addSSInitLater(_doMergeCmd, sheet, data, token);
		},
		columnSize:  _size = function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			
			var data = v[2];
			if (sheet._initiated)
				_doSizeCmd(sheet, data);
			else 
				sheet.addSSInitLater(_doSizeCmd, sheet, data);
		},
		rowSize: _size,
		rowBegin: null,
		rowEnd: null,
		rowOuter: null,
		colBegin: null,
		colEnd: null,
		cellOuter: null,
		cellInner: null,
		celltext: null,
		edittext: null,
		topHeaderOuter: null,
		topHeaderInner: null,
		leftHeaderOuter: null,
		leftHeaderInner: null,
		topHeaderHiddens: null,
		leftHeaderHiddens: null,
		dataPanel: null,
		/**
		 * Sets the customized titles of column header.
		 * @param string array
		 */
		/**
		 * Gets the customized titles of column header.
		 * @return string array
		 */
		columntitle: null,
		/**
		 * Sets the customized titles of row header.
		 * @param string array
		 */
		/**
		 * Gets the customized titles of row header.
		 * @return string array
		 */
		rowtitle: null,
		/**
		 * Sets the default row height of the selected sheet
		 * @param rowHeight the row height
		 */
		/**
		 * Gets the default row height of the selected sheet
		 * @return default value depends on selected sheet
		 */
		/**
		 * TODO avoid use invalidate
		 */
		rowHeight: null,
		/**
		 * Sets the default row height of the selected sheet
		 * @param rowHeight the row height
		 */
		/**
		 * Gets the default column width of the selected sheet
		 * @return default value depends on selected sheet
		 */
		/**
		 * TODO avoid use invalidate
		 */
		columnWidth: null,
		/**
		 * TODO there is no lineh ?
		 */
		lineh: null,
		/**
		 * Sets the top head panel height, must large then 0.
		 * @param topHeight top header height
		 */
		/**
		 * Gets the top head panel height
		 * default 20
		 * @return int
		 */
		/**
		 * TODO avoid use invalidate
		 */
		topPanelHeight: null,
		/**
		 * Sets the left head panel width, must large then 0.
		 * @param leftWidth leaf header width
		 */
		/**
		 * Gets the left head panel width
		 * @return default value is 28
		 */
		/**
		 * TODO avoid use invalidate
		 */
		leftPanelWidth: null,
		/**
		 * cell padding of each cell and header, both on left and right side.
		 */
		cellPadding: null,
		/** 
		 * the encoded URL for the dynamic generated content, or empty
		 */
		scss: null,
		sheetId: null,
		focusRect: null,
		selectionRect: null,
		highLightRect: null,
		mergeRange: null,
		csc: null,
		csr: null,
		//override
		width: function (v) {
			var sheet = this.sheetCtrl;
			if (sheet)
				sheet.resizeTo(v, null);
		},
		//override
		height: function (v) {
			var sheet = this.sheetCtrl;
			if (sheet)
				sheet.resizeTo(null, v);
		},
		/**
		 * Sets true to hide the row head of this spread sheet.
		 * @param boolean v true to hide the row head of this spread sheet.
		 */
		/**
		 * Returns if row head is hidden
		 * @return boolean
		 */
		/**
		 * TODO avoid use invalidate
		 */
		rowHeadHidden: null,
		/**
		 * Sets true to hide the column head of  this spread sheet.
		 * @param boolean v true to hide the row head of this spread sheet.
		 */
		/**
		 * Returns if column head is hidden
		 * @return boolean
		 */
		/**
		 * TODO avoid use invalidate
		 */
		columnHeadHidden: null
	},
	/**
	 * Synchronize widgets position to cell
	 * @param int row row index
	 * @param int col column index
	 */
	syncWidgetPos: function (row, col) {
		var wgtPanel = this.$n('wp'),
			widgets = jq(this.$n('wp')).children('.zswidget'),
			sheet = this.sheetCtrl,
			size = widgets.length;
		while (size--) {
			var n = widgets[size],
				wgt = zk.Widget.$(n.id);
			if (wgt && ((row >= 0 && wgt.getRow() >= row) || (col >= 0 && wgt.getCol() >= col)))
				wgt.adjustLocation();
		}
	},
	/**
	 * Sets the maximum column number of this spread sheet.
	 * for example, if you set to 40, which means it allow column 0 to 39. 
	 * the minimal value of maxcols must large than 0;
	 * @param string
	 */
	setMaxColumn: function (v) {
		this._maxColumnData = jq.evalJSON(v);
		this._setMaxColumn(this._maxColumnData);
	},
	_setMaxColumn: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		if (sheet._initiated)
			_doMaxcolumnCmd(sheet, v);
		else
			sheet.addSSInitLater(_doMaxcolumnCmd, sheet, v);
	},
	_initMaxColumn: function () {
		var data = this._maxColumnData
		if (data)
			this._setMaxColumn(data);
	},
	getMaxColumn: function () {
		var data = this._maxColumnData;
		if (!data) return null;
		
		return data.maxcol;
	},
	getColumnFreeze: function () {
		var data = this._maxColumnData;
		if (!data) return null;
		
		return data.colfreeze;
	},
	/**
	 * Sets the maximum row number of this spread sheet.
	 * for example, if you set to 40, which means it allow row 0 to 39.
	 * the minimal value of maxrows must large than 0;
	 * <br/>
	 * Default : 40.
	 * @param string
	 */
	setMaxRow: function (v) {
		this._maxRowData = jq.evalJSON(v);
		this._setMaxRow(this._maxRowData);
	},
	_setMaxRow: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		if (sheet._initiated)
			_doMaxrowCmd(sheet, v);
		else
			sheet.addSSInitLater(_doMaxrowCmd, sheet, v);
	},
	_initMaxRow: function () {
		var data = this._maxRowData;
		if (data)
			this._setMaxRow(data);
	},
	/**
	 * Returns the maximum row number of this spreadsheet.
	 * @return the maximum row number.
	 */
	getMaxRow: function () {
		var data = this._maxRowData;
		if (!data) return null;
		
		return data.maxrow;
	},
	getRowFreeze: function () {
		var data = this._maxRowData;
		if (!data) return null;
		
		return data.rowfreeze;
	},
	/**
	 * Returns the current focus of spreadsheet to Server
	 */
	setRetrieveFocus: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		sheet._cmdRetriveFocus(jq.evalJSON(v));
	},
	/**
	 * Sets the current focus of spreadsheet
	 */
	setCellFocus: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;

		sheet._cmdCellFocus(jq.evalJSON(v));
	},
	/**
	 * Retrieve client side spreadsheet focus.The cell focus and selection will
	 * keep at last status. It is useful if you want get focus back to
	 * spreadsheet after do some outside processing, for example after user
	 * click a outside button or menu item.
	 * 
	 * @param boolean trigger, true will fire a focus event, false won't.
	 */
	focus: function (trigger) {
		if (zk.ie) {
			var self = this;
			setTimeout(function () {
				self.sheetCtrl.dp.gainFocus(trigger);
			}, 0);
		} else
			this.sheetCtrl.dp.gainFocus(trigger);
	},
	/**
	 * Add editor focus
	 */
	addEditorFocus: function (name, color, row, col) {
		this.sheetCtrl.addEditorFocus(name, color);
	},
	/**
	 * Move the editor focus 
	 */
	moveEditorFocus: function (name, color, row, col) {
		this.sheetCtrl.moveEditorFocus(name, color, zk.parseInt(row), zk.parseInt(col));
	},
	/**
	 * Remove the editor focus
	 */
	removeEditorFocus: function (name, color, row, col) {
		this.sheetCtrl.removeEditorFocus(name);
	},
	/**
	 * Sets the highlight rectangle or sets a null value to hide it.
	 */
	setSelectionHighlight: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		sheet._cmdHighlight(jq.evalJSON(v));
	},
	/**
	 * Sets whether display gridlines.
	 * @param boolean show true to show the gridlines; default is true.
	 */
	setDisplayGridlines: function (show) {
		var sheet = this.sheetCtrl;
		if (!sheet) {
			this._dpGridlines = show; //reserver value for init
			return;
		}
		if (show == !this._hideGridlines) 
			return;

		sheet._cmdGridlines(show);
		this._hideGridlines = !show;
	},
	/**
	 * Returns whether display gridlines.
	 * @return boolean
	 */
	isDisplayGridlines: function () {
		return !this._hideGridlines;
	},
	/**
	 * Sets the selection rectangle of the spreadsheet
	 */
	setSelection: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		sheet._cmdSelection(jq.evalJSON(v));
	},
	/**
	 * Returns whether if the server has registers Cell click event or not
	 * @return boolean
	 */
	_isFireCellEvt: function (type) {
		if (type == "lc" || type == "rc" || type == "dbc") {
			var opt = {asapOnly: true};
			return this.isListen('onCellClick', opt) || this.isListen('onCellRightClick', opt) || this.isListen('onCellDoubleClick', opt);
		}
		return false;
	},
	/**
	 * Returns whether if the server has registers Header click event or not
	 * @return boolean
	 */
	_isFireHeaderEvt: function (type) {
		if (type == "lc" || type == "rc" || type == "dbc") {
			var opt = {asapOnly: true};
			return this.isListen('onHeaderClick', opt) || this.isListen('onHeaderRightClick', opt) || this.isListen('onHeaderDoubleClick', opt);
		}
		return false;
	},
	/**
	 * Fire Header click event
	 * <p> Fires header event to server only if the server registers Header click event
	 * <p> Fires header event at client side
	 * @param string type, the type of the header event, "lc" for left click, "rc" for right click, "dbc" for double click
	 * 
	 */
	fireHeaderEvt: function (type, shx, shy, mousemeta, row, col, mx, my) {
		var sheetId = this.getSheetId(),
			prop = {type: type, shx: shx, shy: shy, key: mousemeta, sheetId: sheetId, row: row, col: col, mx: mx, my: my};
		if (this._isFireHeaderEvt(type)) {
			//1995689 selection rectangle error when listen onCellClick, 
			//use timeout to delay mouse click after mouse up(selection)
			var self = this;
			setTimeout(function() {
				self.fire('onZSSHeaderMouse', prop, {toServer: true});
			}, 0);
		}
		var evtName;
		switch (type) {
		case 'lc': 
			evtName = 'onHeaderClick';
			break;
		case 'rc':
			evtName = 'onHeaderRightClick';
			break;
		case 'dbc':
			evtName = 'onHeaderDoubleClick';
			break;
		}
		if (evtName) {
			var e = new zk.Event(this, evtName, prop);
			e.auStopped = true;
			this.fireX(e);
		}
	},
	/**
	 * Fires Cell Event
	 * <p> 
	 * @param string type
	 */
	fireCellEvt: function (type, shx, shy, mousemeta, row, col, mx, my) {
		var sheetId = this.getSheetId(),
			prop = {type: type, shx: shx, shy: shy, key: mousemeta, sheetId: sheetId, row: row, col: col, mx: mx, my: my};
		if (this._isFireCellEvt(type)) {
			//1995689 selection rectangle error when listen onCellClick, 
			//use timeout to delay mouse click after mouse up(selection)
			var self = this;
			setTimeout(function() {
				self.fire('onZSSCellMouse',	prop, {toServer: true}, 25);
			}, 0);
		}
		var evtName;
		switch (type) {
		case 'lc':
			evtName = 'onCellClick';
			break;
		case 'rc':
			evtName = 'onCellRightClick';
			break;
		case 'dbc':
			evtName = 'onCellDoubleClick';
			break;
		}
		if (evtName) {
			var e = new zk.Event(this, evtName, prop);
			e.auStopped = true;
			this.fireX(e);
		}
	},
	_getTopHeaderFontSize: function () {
		var head = this.$n('tophead'),
			col = head != null ? jq(head).children('div:first')[0] : null;
		if (col && jq(col).attr('zs.t') == 'STheader')
			return jq(col).css('font-size');
		return null;
	},
	_initFrozenArea: function () {
		var rowFreeze = this.getRowFreeze(),
			colFreeze = this.getColumnFreeze();
		if ((rowFreeze && rowFreeze > -1) || (colFreeze && colFreeze > -1)) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			
			if (sheet._initiated)
				syncFrozenArea(sheet);
			else
				sheet.addSSInitLater(syncFrozenArea, sheet);
		}	
	},
	_initControl: function () {
		var sheet = this.sheetCtrl = new zss.SSheetCtrl(this.$n(), this);
		this._loadCSSDirect(this._scss, this.uuid + "-sheet");
		
		this._initMaxColumn();
		this._initMaxRow();
		this._initFrozenArea();

		var show = this._dpGridlines;
		if (typeof show != 'undefined')
			this.setDisplayGridlines(show);
	},
	bind_: function () {
		this.$supers('bind_', arguments);

		this._initControl();
		//zWatch.listen({onShow: this, onHide: this, onSize: this});
		zWatch.listen({onSize: this});
	},
	unbind_: function () {
		//zWatch.unlisten({onShow: this, onHide: this, onSize: this});
		zWatch.unlisten({onSize: this});

		this._maxColumnMap = this._maxRowMap = null;
		
		var n = this.$n(),
			ctrl = n.ctrl;
		if (n.ctrl) {
			ctrl.cleanup();
			removeSSheet(ctrl.sheetid + "-sheet");
			ctrl = null;
		}
		if (window.CollectGarbage)
			window.CollectGarbage();
		
		this.$supers('unbind_', arguments);
	},/*
	onShow: function () {
		//zssd("onVisi" + this.$n());
	},
	onHide: function () {
		//zssd("onHide:" + this.$n());
	},*/
	onSize: function () {
		var sheet = this.sheetCtrl;
		if (!sheet || !sheet._initiated) return;

		sheet._resize();
	},
	domClass_: function (no) {
		return 'zssheet';
	},
	/**
	 * To do, don't use html attr. replace with widget function
	 */
	domAttrs_: function (no) {
		var attrs = this.$supers('domAttrs_', arguments) + ' z.type="zss.ss.SSheet" zs.t="SSheet"',
			rowHeight = this._rowHeight,
			colWidth = this._columnWidth,
			toph = this._topPanelHeight,
			leftw = this._leftPanelWidth,
			cellpad = this._cellPadding,
			css = this._scss,
			maxCol = this.getMaxColumn(),
			maxRow = this.getMaxRow(),
			id = this._sheetId,
			focusRect = this._focusRect,
			selRect = this._selectionRect,
			highLightRect = this._highLightRect,
			merge = this._mergeRange,
			colFreeze = this.getColumnFreeze(),
			rowFreeze = this.getRowFreeze();
		
		if (rowHeight)
			attrs += ' z.rowh="' + rowHeight + '"';
		if (colWidth)
			attrs += ' z.colw="' + colWidth + '"';
		if (toph)
			attrs += ' z.toph="' + toph + '"';
		if (leftw)
			attrs += ' z.leftw="' + leftw + '"';
		if (cellpad)
			attrs += ' z.cellpad="' + cellpad + '"';
		if (css)
			attrs += ' z.scss="' + css + '"';
		if (maxCol)
			attrs += ' z.maxc="' + maxCol + '"';
		if (maxRow)
			attrs += ' z.maxr="' + maxRow + '"';
		if (id)
			attrs += ' z.sheetId="' + id + '"';
		if (focusRect)
			attrs += ' z.fs="' + focusRect + '"';
		if (selRect)
			attrs += ' z.sel="' + selRect + '"';
		if (highLightRect)
			attrs += ' z.hl="' + highLightRect + '"';
		if (this._csc)
			attrs += ' z.csc="' + this._csc + '"';
		if (this._csr)
			attrs += ' z.csr="' + this._csr + '"';
		if (merge)
			attrs += ' z.mers="' + merge + '"';
		if (colFreeze)
			attrs += ' z.fzc="' + colFreeze + '"';
		if (rowFreeze)
			attrs += ' z.fzr="' + rowFreeze + '"';
		return attrs;
	},
	_doDataPanelBlur: function (evt) {
		var sheet = this.sheetCtrl;
		if (sheet.innerClicking <= 0 && sheet.state == zss.SSheetCtrl.FOCUSED) {
			sheet.dp._doFocusLost();
		} else if(sheet.state == zss.SSheetCtrl.FOCUSED) {
			//retrive focus back to focustag
			sheet.dp.gainFocus(true);
		}
	},
	_doDataPanelFocus: function (evt) {
		var sheet = this.sheetCtrl;
		if (sheet.state < zss.SSheetCtrl.FOCUSED)
			sheet.dp.gainFocus(false);
	},
	_doScrollPanelScrolling: function (evt) {
		var sp = this.sheetCtrl.sp;
		if (sp)
			sp._doScrolling(evt);
	},
	_doScrollPanelMouseDown: function (evt) {
		var sp = this.sheetCtrl.sp;
		if (sp)
			sp._doMousedown(evt);
	},
	_doEditboxBlur: function (evt) {
		var sheet = this.sheetCtrl,
			dp =  sheet.dp;
		if (dp)
			dp.stopEditing(sheet.innerClicking > 0 ? "refocus" : "lostfocus");
	},
	_doEditboxKeyPress: function (evt) {
		var editor = this.sheetCtrl.editor;
		if (editor)
			editor.autoAdjust(evt);
	},
	_doEditboxKeyDown: function (evt) {
		var editor = this.sheetCtrl.editor;
		if (editor) {
			if (editor.disabled)
				evt.stop();
			else
				editor._doKeydown(evt);
		}
	},
	_doEditboxKeyUp: function (evt) {
		var editor = this.sheetCtrl.editor;
		if (editor)
			editor._doKeyup(evt);
	},
	_doLeftPanelMouseOver: function (evt) {
		var lp = this.sheetCtrl.lp;
		if (lp)
			lp._doMouseover(evt);
	},
	_doLeftPanelMouseOut: function (evt) {
		var lp = this.sheetCtrl.lp;
		if (lp)
			lp._doMouseout(evt);
	},
	_doTopPanelMouseOver: function (evt) {
		var tp = this.sheetCtrl.tp;
		if (tp)
			tp._doMouseover(evt);
	},
	_doTopPanelMouseOut: function (evt) {
		var tp = this.sheetCtrl.tp;
		if (tp)
			tp._doMouseout(evt);
	},
	_doSelAreaMouseMove: function (evt) {
		var sel = this.sheetCtrl.selArea;
		if (sel)
			sel._doMouseMove(evt);
	},
	_doSelAreaMouseOut: function (evt) {
		var sel = this.sheetCtrl.selArea;
		if (sel)
			sel._doMouseOut(evt);
	},
	doClick_: function (evt) {
		this.sheetCtrl._doMouseleftclick(evt);
		this.$supers('doClick_', arguments);
	},
	/**
	 * override
	 * 
	 * this method will check the current focus information first.
	 * 
	 * Case 1: no focus currently, sets the focus of the spreadsheet to last focus position.
	 * Case 2: spreadsheet has focus, depends on the target, execute the relative mouse down behavior		
	 */
	doMouseDown_: function (evt) {
		this.sheetCtrl._doMousedown(evt);
		this.$supers('doMouseDown_', arguments);
	},
	doMouseUp_: function (evt) {
		this.sheetCtrl._doMouseup(evt);
		this.$supers('doMouseUp_', arguments);
	},
	doRightClick_: function (evt) {
		this.sheetCtrl._doMouserightclick(evt);
		this.$supers('doRightClick_', arguments);
	},
	doDoubleClick_: function (evt) {
		this.sheetCtrl._doMousedblclick(evt);
		this.$supers('doDoubleClick_', arguments);
	},
	_doDragMouseUp: function (evt) {
		var dragHandler = this.sheetCtrl.dragging;
		if (dragHandler)
			dragHandler.doMouseup(evt);
	},
	_doDragMouseDown: function (evt) {
		var dragHandler = this.sheetCtrl.dragging;
		if (dragHandler)
			dragHandler.doMousemove(evt);
	},
	doKeyDown_: function (evt) {
		this.sheetCtrl._doKeydown(evt);
		this.$supers('doKeyDown_', arguments);
	},
	doKeyPress_: function (evt) {
		this.sheetCtrl._doKeypress(evt);
		this.$supers('doKeyPress_', arguments);
	},
	_loadCSSDirect: function (uri, id) {
		var e = document.createElement("LINK");
		if (id) e.id = id;
		e.rel = "stylesheet";
		e.type = "text/css";
		e.href = uri;
		document.getElementsByTagName("HEAD")[0].appendChild(e);
	},
	linkTo: function (href, type, evt) {
		//1: LINK_URL
		//2: LINK_DOCUMENT
		//3: LINK_EMAIL
		//4: LINK_FILE
		if (type == 1 && !evt.ctrlKey) //LINK_URL, no CTRL
			location.href = href;
		else if (type == 1 && evt.ctrlKey)//LINK_URL, with CTRL
			window.open(href);
		else if (type == 3) //LINK_EMAIL
			location.href = href;
//		else if (type == 4) //LINK_FILE
			//TODO LINK_FILE
//		else if (type == 2) //LINK_DOCUMENT
			//TODO LINK_DOCUMENT
	}
}, {
	initLaterAfterCssReady: function (sheet) {
		if (zcss.findRule(".zs_indicator", sheet.id + "-sheet") != null) {
			sheet._initiated = true;

			_calScrollWidth(sheet.comp);
			//since some initial depends on width or height,
			//so first ss initial later must invoke after css ready,
			
			//_doSSInitLater must invoke before loadForVisible,
			//because some urgent initial must invoke first, or loadForVisible will load a wrong block
			sheet._doSSInitLater();
			if (sheet._initmove) {//#1995031
				sheet.showMask(false);
			} else if(!sheet.activeBlock.loadForVisible()){
				//if no loadfor visible send after init, then we must sync the block size
				sheet.sendSyncblock(true);
				sheet.showMask(false);
			}
			if (zk.opera)
				//opera cannot insert rule on special index,
				//so I must create another style sheet to control style rule priority
				createSSheet("", sheet.id + "-sheet-opera");//heigher
			
			//force IE to update CSS
			if (zk.ie6_ || zk.ie7_)
				jq(sheet.activeBlock.comp).toggleClass('zssnosuch');
		} else {
			sheet._initcount++;
			if (sheet._initcount > 20)
				sheet._initcount = 0;
			
			setTimeout(function () {
				zss.Spreadsheet.initLaterAfterCssReady(sheet);
			}, 100);
		}		
	}
});
})();