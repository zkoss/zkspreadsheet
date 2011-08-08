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
	/**
	 * Includes flexbox support function
	 * 
	 * Modernizr v1.7
	 * http://www.modernizr.com
	 *
	 * Developed by: 
	 * - Faruk Ates  http://farukat.es/
	 * - Paul Irish  http://paulirish.com/
	 *
	 * Copyright (c) 2009-2011
	 * Dual-licensed under the BSD or MIT licenses.
	 * http://www.modernizr.com/license/
	 * 
	 */
	function _flexSupport() {
		var docElement = document.documentElement,
			prefixes = ' -webkit- -moz- -o- -ms- -khtml- '.split(' ');
        /**
         * set_prefixed_value_css sets the property of a specified element
         * adding vendor prefixes to the VALUE of the property.
         * @param {Element} element
         * @param {string} property The property name. This will not be prefixed.
         * @param {string} value The value of the property. This WILL be prefixed.
         * @param {string=} extra Additional CSS to append unmodified to the end of
         * the CSS string.
         */
        function set_prefixed_value_css(element, property, value, extra) {
            property += ':';
            element.style.cssText = (property + prefixes.join(value + ';' + property)).slice(0, -property.length) + (extra || '');
        }

        /**
         * set_prefixed_property_css sets the property of a specified element
         * adding vendor prefixes to the NAME of the property.
         * @param {Element} element
         * @param {string} property The property name. This WILL be prefixed.
         * @param {string} value The value of the property. This will not be prefixed.
         * @param {string=} extra Additional CSS to append unmodified to the end of
         * the CSS string.
         */
        function set_prefixed_property_css(element, property, value, extra) {
            element.style.cssText = prefixes.join(property + ':' + value + ';') + (extra || '');
        }
        
        var c = document.createElement('div'),
	        elem = document.createElement('div');
	
	    set_prefixed_value_css(c, 'display', 'box', 'width:42px;padding:0;');
	    set_prefixed_property_css(elem, 'box-flex', '1', 'width:10px;');
	
	    c.appendChild(elem);
	    docElement.appendChild(c);
	
	    var ret = elem.offsetWidth === 42;
	
	    c.removeChild(elem);
	    docElement.removeChild(c);
	
	    return ret;
	}
	function _calScrollWidth () {
		if(zkS.t(zss.Spreadsheet.scrollWidth)) return;
	    // scroll scrolling div
		var body = document.body,
			scr = jq('<div style="position:absolute;top:0px;left:0px;width:50px;height:50px;overflow:auto;"></div>')[0],
			inn = jq('<div style="width:100px;height:100px;"></div>')[0]; //scroll content div
	    // Put the scrolli div in the scrolling div
	    scr.appendChild(inn);

	    // Append the scrolling div to the doc
	    body.appendChild(scr);

	    //calcuate the scrollWidth;	
		zss.Spreadsheet.scrollWidth = (scr.offsetWidth - scr.clientWidth);
		if (zss.Spreadsheet.scrollWidth == 0)
			zss.Spreadsheet.scrollWidth = scr.offsetHeight - scr.clientHeight;
		body.removeChild(scr);
		return;
	}
	
	function _doCellUpdateCmd (sheet, data, token) {
		data = jq.evalJSON(data);
		if (jq.isArray(data)) {
			var i = data.length;
			while (i--)
				sheet._cmdCellUpdate(data[i]);
		} else
			sheet._cmdCellUpdate(data);
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
    		styleObj = document.createElement("style"),
    		sheetobj;
	    jq(styleObj).attr("type", "text/css");
		
	    if (id)
	       jq(styleObj).attr("id", id);
	    if (zk.ie) {
	        head.appendChild(styleObj);
	        sheetobj = styleObj.styleSheet;
	        sheetobj.cssText = cssText;
	    } else {
	        try {
	        	styleObj.appendChild(document.createTextNode(cssText));
	        } catch (e) {
	        	styleObj.cssText = cssText;
	        }
	        head.appendChild(styleObj);
			sheetobj = _getElementSheet(styleObj);
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
	 * get stylesheet object from DOMElement
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
 * Spreadsheet is a is a rich ZK Component to handle EXCEL like behavior
 */
zss.Spreadsheet = zk.$extends(zul.Widget, {
	/**
	 * Indicate whether shall process cell's overflow or not
	 */
	//_shallProcessOverflow: 0
	/**
	 * Process overflow column index
	 */
	_processOverflowCol: 0,
	/**
	 * Map source command and relate column index. For process overflow
	 */
	_srcCmd: {},
	$init: function () {
		this.$supers('$init', arguments);
		this._cssFlex = _flexSupport();
	},
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
		 * Inserts new row or column
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
		 * Removes row or column
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
		/**
		 * Sets sheet protection
		 * <p>
		 * 	Default is false
		 * </p>
		 * @param boolean
		 */
		/**
		 * Returns whether protection is enabled or disabled
		 * @return boolean
		 */
		protect: null,
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
		scss: function (url) {
			zk.loadCSS(this._scss, this.uuid + "-sheet");
		},
		sheetId: null,
		focusRect: null,
		selectionRect: null,
		highLightRect: null,
		mergeRange: null,
		autoFilter: null,
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
		columnHeadHidden: null,
		copysrc: null //flag to show whether a copy source has set
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
				if (self.sheetCtrl)
					self.sheetCtrl.dp.gainFocus(trigger);
			}, 0);
		} else if (this.sheetCtrl)
			this.sheetCtrl.dp.gainFocus(trigger);
	},
	/**
	 * Add editor focus
	 */
	addEditorFocus: function (name, color, row, col) {
		if (this.sheetCtrl)
			this.sheetCtrl.addEditorFocus(name, color);
	},
	/**
	 * Move the editor focus 
	 */
	moveEditorFocus: function (name, color, row, col) {
		if (this.sheetCtrl)
			this.sheetCtrl.moveEditorFocus(name, color, zk.parseInt(row), zk.parseInt(col));
	},
	/**
	 * Remove the editor focus
	 */
	removeEditorFocus: function (name, color, row, col) {
		if (this.sheetCtrl)
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
		//Note. for IE, need to delay set gridline after css ready
		if (sheet._initiated)
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
		var evtnm = zss.Spreadsheet.CELL_MOUSE_EVENT_NAME[type];
		return evtnm && this.isListen(evtnm, {asapOnly: true});
	},
	/**
	 * Returns whether if the server has registers Header click event or not
	 * @return boolean
	 */
	_isFireHeaderEvt: function (type) {
		var evtnm = zss.Spreadsheet.HEADER_MOUSE_EVENT_NAME[type];
		return evtnm && this.isListen(evtnm, {asapOnly: true});
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
		var evtName = zss.Spreadsheet.HEADER_MOUSE_EVENT_NAME[type];
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
	fireCellEvt: function (type, shx, shy, mousemeta, row, col, mx, my, field) {
		var sheetId = this.getSheetId(),
			prop = {type: type, shx: shx, shy: shy, key: mousemeta, sheetId: sheetId, row: row, col: col, mx: mx, my: my};
		if (field)
			prop.field = field;
		if (this._isFireCellEvt(type)) {
			//1995689 selection rectangle error when listen onCellClick, 
			//use timeout to delay mouse click after mouse up(selection)
			var self = this;
			setTimeout(function() {
				self.fire('onZSSCellMouse',	prop, {toServer: true}, 25);
			}, 0);
		}
		var evtName = zss.Spreadsheet.CELL_MOUSE_EVENT_NAME[type];
		if (evtName) {
			var e = new zk.Event(this, evtName, prop);
			e.auStopped = true;
			this.fireX(e);
		}
	},
	_getTopHeaderFontSize: function () {
		var head = this.$n('tophead'),
			col = head != null ? head.firstChild : null;
		if (col && col.getAttribute('zs.t') == 'STheader')
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
		if (this.getSheetId() == null) //no sheet at all
			return;
		var sheet = this.sheetCtrl = new zss.SSheetCtrl(this.$n(), this);
		
		this._initMaxColumn();
		this._initMaxRow();
		this._initFrozenArea();

		var show = this._dpGridlines;
		if (typeof show != 'undefined')
			this.setDisplayGridlines(show);
	},
	bind_: function () {
		_calScrollWidth();
		this.$supers('bind_', arguments);

		this._initControl();
		zWatch.listen({onShow: this, onSize: this, onResponse: this});
	},
	unbind_: function () {
		zWatch.unlisten({onShow: this, onSize: this, onResponse: this});

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
	},
	/**
	 * Set process cell's overflow column range
	 */
	setProcessOverflowCol_: function (colIdx, cmd) {
		this._shallProcessOverflow = true;
		this._processOverflowCol = Math.max(colIdx, this._processOverflowCol);
		if (cmd != undefined)
			this._srcCmd[cmd] = colIdx;
	},
	onResponse: function () {
		var col = this._processOverflowCol,
			overflow = this._shallProcessOverflow,
			cmd = this._srcCmd;

		if (overflow) {
			var sCtrl = this.sheetCtrl,
				cBlock = sCtrl.cp.block,
				tBlock = sCtrl.tp.block,
				lBlock = sCtrl.lp.block;
			sCtrl.activeBlock._processTextOverflow(col, cmd);
			if (cBlock)
				cBlock._processTextOverflow(col, cmd);
			if (tBlock)
				tBlock._processTextOverflow(col, cmd);
			if (lBlock)
				lBlock._processTextOverflow(col, cmd);
			this._shallProcessOverflow = false;
			this._processOverflowCol = 0;
			this._srcCmd = {};
		}
	},
	onShow: _zkf = function () {
		var sheet = this.sheetCtrl;
		if (!sheet || !sheet._initiated) return;

		sheet._resize();
	},/*
	onHide: function () {
		//zssd("onHide:" + this.$n());
	},*/
	onSize: _zkf,
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
		if (this.sheetCtrl)
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
		if (this.sheetCtrl)
			this.sheetCtrl._doMousedown(evt);
		this.$supers('doMouseDown_', arguments);
	},
	doMouseUp_: function (evt) {
		if (this.sheetCtrl)
			this.sheetCtrl._doMouseup(evt);
		this.$supers('doMouseUp_', arguments);
	},
	doRightClick_: function (evt) {
		if (this.sheetCtrl)
			this.sheetCtrl._doMouserightclick(evt);
		this.$supers('doRightClick_', arguments);
	},
	doDoubleClick_: function (evt) {
		if (this.sheetCtrl)
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
		if (this.sheetCtrl)
			this.sheetCtrl._doKeydown(evt);
		this.$supers('doKeyDown_', arguments);
	},
	afterKeyDown_: function (evt) {
		if (this.sheetCtrl.state != zss.SSheetCtrl.EDITING) {
			this.$supers('afterKeyDown_', arguments);
			//feature #26: Support copy/paste value to local Excel
			var keyCode = evt.keyCode;
			if (this.isListen('onCtrlKey', {any:true}) && 
				(keyCode == 67 || keyCode == 86)) {
				var parsed = this._parsedCtlKeys,
					ctrlKey = evt.ctrlKey ? 1: evt.altKey ? 2: evt.shiftKey ? 3: 0;
				if (parsed && 
					parsed[ctrlKey][keyCode]) {
					//Widget.js will stop event, if onCtrlKey reg ctrl + c and ctrl + v. restart the event
					evt.domStopped = false;
				}
			}
		}
		//avoid onCtrlKey to be eat in editing mode.
	},
	doKeyPress_: function (evt) {
		if (this.sheetCtrl)
			this.sheetCtrl._doKeypress(evt);
		this.$supers('doKeyPress_', arguments);
	},
	//feature#161: Support copy&paste from clipboard to a cell
	doKeyUp_: function (evt) {
		this.$supers('doKeyUp_', arguments);
		if (this.sheetCtrl && this.sheetCtrl.state == zss.SSheetCtrl.FOCUSED)
			this.sheetCtrl._doKeyup(evt);
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
	CELL_MOUSE_EVENT_NAME: {lc:'onCellClick', rc:'onCellRightClick', dbc:'onCellDoubleClick', af:'onFilter'},
	HEADER_MOUSE_EVENT_NAME: {lc:'onHeaderClick', rc:'onHeaderRightClick', dbc:'onHeaderDoubleClick'},
	SRC_CMD_SET_COL_WIDTH: 'setColWidth',
	initLaterAfterCssReady: function (sheet) {
		if (zcss.findRule(".zs_indicator", sheet.id + "-sheet") != null) {

			sheet._cmdGridlines(sheet._wgt.isDisplayGridlines());
			sheet._initiated = true;
			//since some initial depends on width or height,
			//so first ss initial later must invoke after css ready,
			
			//_doSSInitLater must invoke before loadForVisible,
			//because some urgent initial must invoke first, or loadForVisible will load a wrong block
			sheet._doSSInitLater();
			if (sheet._initmove) {//#1995031
				sheet.showMask(false);
			} else if(zk(sheet._wgt).isRealVisible() && !sheet.activeBlock.loadForVisible()){
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
			setTimeout(function () {
				zss.Spreadsheet.initLaterAfterCssReady(sheet);
			}, 10);
		}		
	}
});
})();