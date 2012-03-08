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
	
	function doUpdate(wgt, data, token) {
		wgt.sheetCtrl._skipMove = data.sk; //whether to skip moving the focus/selection after update
		wgt.sheetCtrl._cmdCellUpdate(data);
		if (token)
			zkS.doCallback(token);
		delete wgt.sheetCtrl._skipMove; //reset to don't skip
	}
	
	function copyRow(lCol, rCol, srcRow) {
		var row = {
			r: srcRow.r,
			heightId: srcRow.heightId,
			cells: {},
			update: srcRow.update,
			getCell: srcRow.getCell
		}
		copyCells(lCol, rCol, srcRow, row);
		return row;
	}
	
	function copyCells(lCol, rCol, srcRow, dstRow) {
		var srcCells = srcRow.cells,
			dstCells = dstRow.cells;
		for (var c = lCol; c <= rCol; c++) {
			dstCells[c] = cloneCell(srcCells[c]);
		}
	}
	
	function cloneCell(cell) {
		var c = {};
		for (var p in cell) {
			c[p] = cell[p];
		}
		return c;
	}
	
	/**
	 * Sync frozen range heightId with main range, 
	 * the heightId may exist in client side only when row contains wrap cell
	 */
	function syncHeightId(srcRange, dstRange) {
		var srcRows = srcRange.rows,
			dstRows = dstRange.rows,
			dstRowHeaders = dstRange.rowHeaders;
		for (var rowNum in dstRows) {
			var row = dstRows[rowNum],
				srcRow = srcRows[rowNum];
			if (srcRow) {
				var hId = srcRow.heightId;
				if (hId && !row.heightId) {
					dstRange.updateRowHeightId(rowNum, hId);	
				}
			}
		}
	}

	function doBlockUpdate(wgt, json, token) {
		var ar = wgt._cacheCtrl.getSelectedSheet(),
			tp = json.type;
		if (ar && tp != 'ack') { //fetch cell will empty return,(not thing need to fetch)
			var d = json.data,
				tRow = json.top,
				lCol = json.left,
				bRow = tRow + json.height - 1,
				rCol = lCol + json.width - 1,
				leftFrozen = json.leftFrozen,
				topFrozen = json.topFrozen;
			
			ar.update(d);
			if (leftFrozen) {
				ar.leftFrozen = new zss.ActiveRange(leftFrozen.data);
				//TODO: fine-tune the following code
				var	srcRows = ar.rows,
					f = ar.leftFrozen,
					fRect = f.rect,
					fRows = f.rows,
					fTop = fRect.top,
					fBtm = fRect.bottom,
					fRight = fRect.right;
				if (ar.rect.top < fTop) {
					//copy top rows
					for (var r = ar.rect.top; r < fTop; r++) {
						var srcRow = srcRows[r];
						if (srcRow) {
							fRows[r] = copyRow(0, fRight, srcRow);
						}
					}
					fRect.top = ar.rect.top;
				}
				if (fBtm < ar.rect.bottom) {
					//copy bottom rows
					for (var r = fRect.bottom + 1; r <= ar.rect.bottom; r++) {
						var srcRow = srcRows[r];
						if (srcRow) {
							fRows[r] = copyRow(0, fRight, srcRow);	
						}
					}
					fRect.bottom = ar.rect.bottom;
				}
				//row contains wrap cell may have height Id on client side
				syncHeightId(ar, ar.leftFrozen);
			}
			if (topFrozen) {
				ar.topFrozen = new zss.ActiveRange(topFrozen.data);
				//TODO: fine-tune the following code
				var left = ar.rect.left,
					right = ar.rect.right,
					srcRows = ar.rows,
					f = ar.topFrozen,
					fRect = f.rect,
					fLeft = fRect.left,
					fRight = fRect.right;
				if (left < fLeft) {
					//copy left cells
					var fRows = f.rows,
						fTop = fRect.top,
						fBtm = fRect.bottom;
					for (r = fTop; r <= fBtm; r++) {
						var srcRow = srcRows[r],
							dstRow = fRows[r];
						if (srcRow) {
							copyCells(left, fLeft - 1, srcRow, dstRow);	
						}
					}
				}
				if (fRight < right) {
					//copy right cells
					var fRows = f.rows,
						fTop = fRect.top,
						fBtm = fRect.bottom;
					for (r = fTop; r <= fBtm; r++) {
						var srcRow = srcRows[r],
							dstRow = fRows[r];
						if (srcRow) {
							copyCells(fRect.right + 1, right, srcRow, dstRow);
						}
					}
				}
				//row contains wrap cell may have height Id on client side
				syncHeightId(ar, ar.topFrozen);
			}
			wgt.sheetCtrl._cmdBlockUpdate(tp, d.dir, tRow, lCol, bRow, rCol, leftFrozen, topFrozen);
		}
		if (token)
			zkS.doCallback(token);
	}
	
	function _doInsertRCCmd (sheet, data, token) {
		sheet._cmdInsertRC(data);
		if (token != "")
			zkS.doCallback(token);	
	}
	function _doRemoveRCCmd (sheet, data, token) {
		sheet._cmdRemoveRC(data, true);
		if (token != "")
			zkS.doCallback(token);
	}
	function _doMergeCmd (sheet, data, token) {
		sheet._cmdMerge(jq.evalJSON(data), true);
		if (token != "")
			zkS.doCallback(token);
	}
	function _doSizeCmd (sheet, data, token) {
		sheet._cmdSize(data, true);
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
	 * Returns CSS link in head
	 * 
	 * @return StyleSheet object
	 */
	function hasCSS (id) {
		var head = document.getElementsByTagName("head")[0];
		for (var n = head.firstChild; n; n = n.nextSibling) {
			if (n.id == id) {
				return n;
			}
		}
		return null;
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
var Spreadsheet = 
zss.Spreadsheet = zk.$extends(zul.wgt.Div, {
	/**
	 * Indicate Ctrl-Paste event key down status
	 */
	_ctrlPasteDown: false,
	/**
	 * Indicate whether to always open hyperlink in a separate browser tab window; default true.
	 * <p>If this value is true, Spreadsheet will always open the link in a separate browser tab window.</p>
	 * <p>If this value is false, Spreadsheet will click to open the link in the same browser tab window; or 
	 * CTRL-click to open the link in a separate browser tab window.</p>
	 * @see #linkTo
	 */
	_linkToNewTab: true, //ZSS-13: Support Open hyperlink in a separate browser tab window
	_cellPadding: 2,
	_protect: false,
	_maxRows: 20,
	_maxColumns: 10,
	_rowFreeze: -1,
	_columnFreeze: -1,
	_clientCacheDisabled: false,
	_topPanelHeight: 20,
	_leftPanelWidth: 36,
	_maxRenderedCellSize: 8000,
	_displayGridlines: true,
	/**
	 * Contains spreadsheet's toolbar
	 */
	//_topPanel: null
	/**
	 * Contains zss.Formulabar, zss.SSheetCtrl
	 */
	//cave: null
	$init: function () {
		this.$supers(Spreadsheet, '$init', arguments);
		
		this.appendChild(this.cave = new zul.layout.Borderlayout({
			vflex: true, sclass: 'zscave'
		}));//contains zss.SSheetCtrl
	},
	$define: {
		/**
		 * Indicate cache data at client, won't prune data while scrolling sheet
		 */
		clientCacheDisabled: null,
		/**
		 * Sets labels
		 */
		labels: function (v) {
			var labelCtrl = this._labelsCtrl = new zss.Labels();
			for (var key in v) {
				//key.substr(4): remove prefix 'zss.';
				var val = v[key],
					postfix = key.substr(4),
					fn = 'set' + postfix.charAt(0).toUpperCase() + postfix.substr(1);
				labelCtrl[fn].call(labelCtrl, val);
			}
		},
		/**
		 * synchronized update data
		 * @param array
		 */
		dataBlockUpdate: _dataUpdate = function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			var token = v[0],
				json = jq.evalJSON(v[2]);
			
			if (sheet._initiated) {
				doBlockUpdate(this, json, token);
			} else {
				sheet.addSSInitLater(doBlockUpdate, this, json, token);
			}
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
			
			if (sheet._initiated) {
				doUpdate(this, data, token);
			} else {
				sheet.addSSInitLater(doUpdate, this, data, token);
			}
		},
		dataUpdateStart: _updateCell,
		dataUpdateCancel: _updateCell,
		dataUpdateStop: _updateCell,
		dataUpdateRetry: _updateCell,
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
			else {
				sheet.addSSInitLater(_doInsertRCCmd, sheet, data, token);
			}
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
			else {
				sheet.addSSInitLater(_doRemoveRCCmd, sheet, data, token);
			}
		},
		mergeCell: function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			var token = v[0], 
				data = v[2];

			if (sheet._initiated)
				_doMergeCmd(sheet, data, token);
			else {
				sheet.addSSInitLater(_doMergeCmd, sheet, data, token);
			}
		},
		columnSize:  _size = function (v) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			
			var data = v[2];
			if (sheet._initiated)
				_doSizeCmd(sheet, data);
			else {
				sheet.addSSInitLater(_doSizeCmd, sheet, data);
			}
		},
		/**
		 * Sets sheet protection. Default is false
		 * @param boolean
		 */
		/**
		 * Returns whether protection is enabled or disabled
		 * @return boolean
		 */
		protect: function (v) {
			var sheetPanel = this._sheetPanel;
			if (sheetPanel) {
				var sheet = this.sheetCtrl;
				if (sheet) {
					sheet.fireProtectSheet(v);
				}
				sheetPanel.getSheetSelector().setProtectSheetCheckmark(v);
			}
		},
		rowSize: _size,
		preloadRowSize: null,
		preloadColumnSize: null,
		initRowSize: null,
		initColumnSize: null,
		maxRenderedCellSize: null,
		/**
		 * Sets the maximum visible number of rows of this spreadsheet. For example, if you set
		 * this parameter to 40, it will allow showing only row 0 to 39. The minimal value of max number of rows
		 * must large than 0; <br/>
		 * Default : 20.
		 * 
		 * @param maxrows  the maximum visible number of rows
		 */
		/**
		 * Returns the maximum visible number of rows of this spreadsheet. You can assign
		 * new number by calling {@link #setMaxrows(int)}.
		 * 
		 * @return the maximum visible number of rows.
		 */
		maxRows: null,
		/**
		 * Sets the maximum column number of this spreadsheet.
		 * for example, if you set to 40, which means it allow column 0 to 39. 
		 * the minimal value of maxcols must large than 0;
		 * <br/>
		 * Default : 10.
		 * 
		 * @param string
		 */
		/**
		 * Returns the maximum visible number of columns of this spreadsheet.
		 * 
		 * @return the maximum visible number of columns 
		 */
		maxColumns: null,
		/**
		 * Sets the row freeze of this spreadsheet
		 * 
		 * @param rowfreeze row index
		 */
		/**
		 * Returns the row freeze index of this spreadsheet, zero base. Default : -1
		 * 
		 * @return the row freeze of selected sheet.
		 */
		rowFreeze: null,
		/**
		 * Sets the column freeze of this spreadsheet
		 * 
		 * @param columnfreeze  column index
		 */
		/**
		 * Returns the column freeze index of this spreadsheet, zero base. Default : -1
		 * 
		 * @return the column freeze of selected sheet.
		 */
		columnFreeze: null,
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
		rowHeight: null,
		/**
		 * Sets the default column width of the selected sheet
		 * @param columnWidth the default column width
		 */
		/**
		 * Gets the default column width of the selected sheet
		 * @return default value depends on selected sheet
		 */
		columnWidth: null,
		/**
		 * Sets the top head panel height, must large then 0.
		 * @param topHeight top header height
		 */
		/**
		 * Gets the top head panel height
		 * default 20
		 * @return int
		 */
		topPanelHeight: null,
		/**
		 * Sets the left head panel width, must large then 0.
		 * @param leftWidth leaf header width
		 */
		/**
		 * Gets the left head panel width
		 * @return default value is 36
		 */
		leftPanelWidth: null,
		/**
		 * cell padding of each cell and header, both on left and right side.
		 */
		cellPadding: null,
		/** 
		 * the encoded URL for the dynamic generated content, or empty
		 * 
		 * @param string href CSS link
		 */
		scss: function (href) {
			var el = this.getCSS();
			if (el) {
				jq(el).attr('href', href);
			}
		},
		/**
		 * Sets selected sheet uuid
		 * 
		 * @param string sheey uuid
		 * @param boolean fromServer
		 * @param zss.Range visible range (from client) 
		 */
		sheetId: function (id, fromServer, visRange) {
			//For isSheetCSSReady() to work correctly.
			//when during select sheet in client side, server send focus au response first (set attributes later), 
			// _sheetId will be last selected sheet, cause isSheetCSSReady() doesn't work correctly 
			this._invalidatedSheetId = false;
			
			var sheetCtrl = this.sheetCtrl,
				cacheCtrl = this._cacheCtrl,
				sheetPanel = this._sheetPanel,
				sheetSelector = sheetPanel ? sheetPanel.getSheetSelector() : null;
			if (sheetSelector)
				sheetSelector.setSelectedSheet(id);
			if (sheetCtrl && cacheCtrl && cacheCtrl.getSelectedSheet().sheetId != id) {
				if (!fromServer) {
					cacheCtrl.setSelectedSheetBy(id);	
				}
				sheetCtrl.doSheetSelected(visRange);
			}
			var loadSheetStart = this._loadSheetStart;
			if (loadSheetStart) {
				this._loadSheetStart = false;
			}
		},
		/**
		 * Sets whether display gridlines.
		 * 
		 * Default: true
		 * @param boolean show true to show the gridlines;
		 */
		/**
		 * Returns whether display gridlines. default is true
		 * 
		 * @return boolean
		 */
		displayGridlines: function (show) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;

			if (this.isSheetCSSReady()) {
				sheet.setDisplayGridlines(show);
			} else {
				//set cell focus after CSS ready
				sheet.addSSInitLater(function () {
					sheet.setDisplayGridlines(show);
				});
			}
		},
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
		rowHeadHidden: null,
		/**
		 * Sets true to hide the column head of  this spread sheet.
		 * @param boolean v true to hide the row head of this spread sheet.
		 */
		/**
		 * Returns if column head is hidden
		 * @return boolean
		 */
		columnHeadHidden: null,
		/**
		 * Sets whether show toolbar or not
		 * 
		 * Default: false
		 * @param boolean show true to show toolbar
		 */
		/**
		 * Returns whther show toolbar
		 * @return boolean
		 */
		showToolbar: function (show) {
			var w = this._toolbar;
			if (!w && show) {
				var tb = this._toolbar = new zss.Toolbar(this),
					topPanel = this.getTopPanel();
				topPanel.appendChild(tb);
				topPanel.setHeight(tb.getSize());
			} else if (w) {
				w.setVisible(show);
			}
		},
		actionDisabled: function (v) {
			var tb = this._toolbar;
			if (tb)
				tb.setDisabled(v);
		},
		/**
		 * Sets whether show formula bar or not
		 * @param boolean show true to show formula bar
		 */
		/**
		 * Returns whether show formula bar
		 * @return boolean
		 */
		showFormulabar: function (show) {
			var w = this._formulabar;
			if (!w && show) {
				this.cave.appendChild(this._formulabar = new zss.Formulabar(this));
			} else if (w) {
				w.setVisible(show);
			}
		},
		/**
		 * Sets whether show sheet panel or not
		 * @param true if want to show sheet tab panel
		 */
		/**
		 * Returns whether show sheet panel
		 * @return boolean
		 */
		showSheetpanel: function (show) {
			var w = this._sheetPanel;
			if (!w && show) {
				this.cave.appendChild(this._sheetPanel = new zss.Sheetpanel(this));
			} else if (w) {
				w.setVisible(show);
			}
		},
		/**
		 * Sets sheet's name and uuid of book
		 */
		sheetLabels: function (v) {
			var sheetPanel = this._sheetPanel;
			if (sheetPanel) {
				sheetPanel.getSheetSelector().setSheetLabels(v);
			}
		},
		copysrc: null //flag to show whether a copy source has set
	},
	getTopPanel: function () {
		var tp = this._topPanel;
		if (!tp) {
			tp = this._topPanel = new zul.layout.Borderlayout({vflex: 'min'});
			this.insertBefore(tp, this.firstChild);
		}
		return tp;
	},
	getSheetCSSId: function () {
		return this.uuid + '-sheet';
	},
	getSelectorPrefix: function () {
		return '#' + this.uuid;
	},
	/**
	 * Returns whether CSS loaded from server or not
	 */
	isSheetCSSReady: function () {
		if (this._invalidatedSheetId) {//set by zss.Sheetpanel, indicate current sheetId is invalidated
			return false;
		}
		return !!zcss.findRule(this.getSelectorPrefix() + " .zs_indicator_" + this.getSheetId(), this.getSheetCSSId());
	},
	/**
	 * Sets active range
	 */
	setActiveRange: function (v) {
		var c = this._cacheCtrl;
		if (!c) {
			this._cacheCtrl = c = new zss.CacheCtrl(this, v);
			
			var center = new zul.layout.Center({border: 0});
			center.appendChild(this.sheetCtrl = new zss.SSheetCtrl(this));
			this.cave.appendChild(center);
		} else {
			var sheet = this.sheetCtrl,
				range;
			if (sheet) {
				c.setSelectedSheet(v);
				this._triggerContentsChanged = true;
			}
		}
	},
	getUpload: function () {
		return this._$upload;
	},
	onChildAdded_: function (child) {
		if (child.$instanceof(zss.Upload)) {
			this._$upload = child;
		}
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
	 * Returns the current focus of spreadsheet to Server
	 */
	setRetrieveFocus: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		if (this.isSheetCSSReady()) {
			sheet._cmdRetriveFocus(v);
		} else {
			sheet.addSSInitLater(function () {
				sheet._cmdRetriveFocus(v);
			});
		}
	},
	/**
	 * Sets the current focus of spreadsheet
	 */
	setCellFocus: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;

		if (this.isSheetCSSReady()) {
			sheet._cmdCellFocus(v);
		} else {
			//set cell focus after CSS ready
			sheet.addSSInitLater(function () {
				sheet._cmdCellFocus(v);
			});
		}
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
				var sht = self.sheetCtrl;
				if (sht && sht._initiated)
					sht.dp.gainFocus(trigger);
			}, 0);
		} else if (this.sheetCtrl)
			this.sheetCtrl.dp.gainFocus(trigger);
	},
	/**
	 * Returns whether child DOM Element has focus or not
	 */
	hasFocus: function () {
		return jq.isAncestor(this.$n(), document.activeElement);
	},
	/**
	 * Add editor focus
	 */
	addEditorFocus: function (id, name, color, row, col) {
		if (this.sheetCtrl)
			this.sheetCtrl.addEditorFocus(id, name, color);
	},
	/**
	 * Move the editor focus 
	 */
	moveEditorFocus: function (id, name, color, row, col) {
		if (this.sheetCtrl)
			this.sheetCtrl.moveEditorFocus(id, name, color, zk.parseInt(row), zk.parseInt(col));
	},
	/**
	 * Remove the editor focus
	 */
	removeEditorFocus: function (id, name, color, row, col) {
		if (this.sheetCtrl)
			this.sheetCtrl.removeEditorFocus(id);
	},
	/**
	 * Sets the highlight rectangle or sets a null value to hide it.
	 */
	setSelectionHighlight: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		if (this.isSheetCSSReady()) {
			sheet._cmdHighlight(v);
		} else {
			//set highlight after CSS ready
			sheet.addSSInitLater(function () {
				sheet._cmdHighlight(v);
			});
		}
	},
	/**
	 * Sets the selection rectangle of the spreadsheet
	 */
	setSelection: function (v) {
		var sheet = this.sheetCtrl;
		if (!sheet) return;
		
		if (this.isSheetCSSReady()) {
			sheet._cmdSelection(v);
		} else {
			//set selection after CSS ready
			sheet.addSSInitLater(function () {
				sheet._cmdSelection(v);
			});
		}
	},
	/**
	 * Returns whether if the server has registers Cell click event or not
	 * @return boolean
	 */
	_isFireCellEvt: function (type) {
		var evtnm = zss.Spreadsheet.CELL_MOUSE_EVENT_NAME[type];
		if ('onFilter' == evtnm) { //server side prepare auto filter popup information for client side
			return true;
		}
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
	/**
	 * Fire Move widget event
	 * 
	 * type
	 * <ul>
	 * 	<li>onWidgetSize</li>
	 * 	<li>onWidgetMove</li>
	 * </li>
	 * 
	 * @param string wgt the widget type
	 * @param string type the event type
	 * @param string id the id of the widget
	 * @param int dx1 the x coordinate within the first cell
	 * @param int dy1 the y coordinate within the first cell
	 * @param int dx2 the x coordinate within the second cell
	 * @param int dy2 the y coordinate within the second cell
	 * @param int col1 the column (0 based) of the first cell
	 * @param int row1 the row (0 based) of the first cell
	 * @param int col2 the column (0 based) of the second cell
	 * @param int row2 the row (0 based) of the second cell
	 */
	fireMoveWidgetEvt: function (wgt, type, id, dx1, dy1, dx2, dy2, col1, row1, col2, row2) {
		this.fire('onZSSMoveWidget', {wgt: wgt, type: type, id: id, dx1: dx1, dy1: dy1, 
			dx2: dx2, dy2: dy2, col1: col1, row1: row1, col2: col2, row2: row2}, {toServer: true}, 25);
	},
	/**
	 * Fire widget control key event
	 * 
	 * @param string wgt the widget type
	 * @param string id
	 * @param int keyCode
	 * @param boolean ctrlKey
	 * @param boolean shiftKey
	 * @param boolean altKey
	 */
	fireWidgetCtrlKeyEvt: function (wgt, id, keyCode, ctrlKey, shiftKey, altKey) {
		this.fire('onZSSWidgetCtrlKey', {wgt: wgt, id: id, keyCode: keyCode, 
			ctrlKey: ctrlKey, shiftKey: shiftKey, altKey: altKey}, {toServer: true}, 25);
	},
	fireToolbarAction: function (act, extra) {
		var data = {sheetId: this.getSheetId(), tag: 'toolbar', act: act};
		this.fire('onZSSAction', zk.copy(data, extra), {toServer: true});
	},
	fireSheetAction: function (act, extra) {
		var data = {sheetId: this.getSheetId(), tag: 'sheet', act: act};
		this.fire('onZSSAction', zk.copy(data, extra), {toServer: true});
	},
	/**
	 * Fetch active range. Currently fetch north/south/west/south direction
	 */
	fetchActiveRange: function (top, left, right, bottom) {
		if (!this._fetchActiveRange) {
			this.fire('onZSSFetchActiveRange', 
				{sheetId: this.getSheetId(), top: top, left: left, right: right, bottom: bottom}, {toServer: true});
			this._fetchActiveRange = true;
		}
	},
	/**
	 * Recive active range data
	 */
	setActiveRangeUpdate: function (v) {
		this._fetchActiveRange = null;
		var cacheCtrl = this._cacheCtrl;
		if (cacheCtrl) {
			cacheCtrl.getSelectedSheet().fetchUpdate(v);
			this.sheetCtrl.sendSyncblock();
		}
	},
	_initFrozenArea: function () {
		var rowFreeze = this.getRowFreeze(),
			colFreeze = this.getColumnFreeze();
		if ((rowFreeze && rowFreeze > -1) || (colFreeze && colFreeze > -1)) {
			var sheet = this.sheetCtrl;
			if (!sheet) return;
			
			if (sheet._initiated)
				syncFrozenArea(sheet);
			else {
				sheet.addSSInitLater(syncFrozenArea, sheet);
			}
		}	
	},
	/**
	 * Returns whether load CSS or not
	 * 
	 * @return boolean
	 */
	getCSS: function () {
		return hasCSS(this.uuid + "-sheet");
	},
	_initControl: function () {
		if (this.getSheetId() == null) //no sheet at all
			return;
		
		var cssId = this.uuid + '-sheet';
		if (!hasCSS(cssId)) { //unbind may remove css, need to check again
			zk.loadCSS(this._scss, cssId);
		}
		
		this._initFrozenArea();
	},
	bind_: function (desktop, skipper, after) {
		_calScrollWidth();
		this._initControl();
		this.$supers('bind_', arguments);
		
		var sheet = this.sheetCtrl;
		if (sheet) {
			sheet.fireProtectSheet(this.isProtect());
			sheet.fireDisplayGridlines(this.isDisplayGridlines());
		}
		
		zWatch.listen({onResponse: this});
	},
	unbind_: function () {
		zWatch.unlisten({onResponse: this});
		
		this._cacheCtrl = this._maxColumnMap = this._maxRowMap = null;
		
		removeSSheet(this.getSheetCSSId());
		this.$supers('unbind_', arguments);
		if (window.CollectGarbage)
			window.CollectGarbage();
	},
	onResponse: function () {
		if (this._triggerContentsChanged != undefined) {
			this.sheetCtrl.fire('onContentsChanged');
			delete this._triggerContentsChanged;
		}
	},
	domClass_: function (no) {
		return 'zssheet';
	},
	_doDataPanelBlur: function (evt) {
		var sheet = this.sheetCtrl;
		if (sheet.innerClicking <= 0 && sheet.state == zss.SSheetCtrl.FOCUSED) {
			//TODO: check zk.currentFocus, if child of spreadsheet, do not _doFocusLost 
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
		var sheet = this.sheetCtrl;
		if (sheet) {
			sheet._doKeydown(evt);
			// CTRL-V: a flag that whether stop event 
			// for avoid multi-paste same clipboard content to focus textarea or not
			if (evt.ctrlKey && evt.keyCode == 86) {
				this._ctrlPasteDown = true;
			}
		}
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
		var sheet = this.sheetCtrl;
		if (sheet && sheet.state == zss.SSheetCtrl.FOCUSED) {
			sheet._doKeyup(evt);
		}
		this._ctrlPasteDown = false;
	},
	linkTo: function (href, type, evt) {
		//1: LINK_URL
		//2: LINK_DOCUMENT
		//3: LINK_EMAIL
		//4: LINK_FILE
		if (type == 1) {
			if (this._linkToNewTab || evt.ctrlKey) //LINK_URL, always link to new tab window or CTRL-click
				window.open(href);
			else //LINK_URL, no CTRL
				location.href = href;
		} else if (type == 3) //LINK_EMAIL
			location.href = href;
//		else if (type == 4) //LINK_FILE
			//TODO LINK_FILE
//		else if (type == 2) //LINK_DOCUMENT
			//TODO LINK_DOCUMENT
	}
}, {
	CELL_MOUSE_EVENT_NAME: {lc:'onCellClick', rc:'onCellRightClick', dbc:'onCellDoubleClick', af:'onFilter', dv:'onValidateDrop'},
	HEADER_MOUSE_EVENT_NAME: {lc:'onHeaderClick', rc:'onHeaderRightClick', dbc:'onHeaderDoubleClick'},
	SRC_CMD_SET_COL_WIDTH: 'setColWidth',
	initLaterAfterCssReady: function (sheet) {
		var wgt = sheet._wgt,
			sheetId = wgt.getSheetId();
		if (wgt.isSheetCSSReady()) {
			sheet._initiated = true;
			//since some initial depends on width or height,
			//so first ss initial later must invoke after css ready,
			
			//_doSSInitLater must invoke before loadForVisible,
			//because some urgent initial must invoke first, or loadForVisible will load a wrong block
			sheet._doSSInitLater();
			if (sheet._initmove) {//#1995031
				sheet.showMask(false);
			} else if(zk(sheet._wgt).isRealVisible() &&	!sheet.activeBlock.loadForVisible()){
				//if no loadfor visible send after init, then we must sync the block size
				sheet.showMask(false);
			}
			if (zk.opera)
				//opera cannot insert rule on special index,
				//so I must create another style sheet to control style rule priority
				createSSheet("", wgt.uuid + "-sheet-opera");//heigher
			
			//force IE to update CSS
			if (zk.ie6_ || zk.ie7_)
				jq(sheet.activeBlock.comp).toggleClass('zssnosuch');
		} else {
			setTimeout(function () {
				zss.Spreadsheet.initLaterAfterCssReady(sheet);
			}, 1);
		}		
	}
});


})();