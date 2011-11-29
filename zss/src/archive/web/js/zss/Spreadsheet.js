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
		wgt.sheetCtrl._cmdCellUpdate(data);
		if (token)
			zkS.doCallback(token);
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

	function doBlockUpdate(wgt, json, token) {
		var ar = wgt._activeRange,
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
				ar.leftFrozen = newCachedRange(leftFrozen.data);
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
						fRows[r] = copyRow(0, fRight, srcRows[r]);
					}
					fRect.top = ar.rect.top;
				}
				if (fBtm < ar.rect.bottom) {
					//copy bottom rows
					for (var r = fRect.bottom + 1; r <= ar.rect.bottom; r++) {
						fRows[r] = copyRow(0, fRight, srcRows[r]);
					}
					fRect.bottom = ar.rect.bottom;
				}
			}
			if (topFrozen) {
				ar.topFrozen = newCachedRange(topFrozen.data);
				var left = ar.rect.left,
					right = ar.rect.right,
					srcRows = ar.rows,
					f = ar.topFrozen,
					fRect = f.rect,
					fLeft = fRect.left,
					fRight = fLeft.right;
				if (left < fLeft) {
					//copy left cells
					var fRows = f.rows,
						fTop = fRect.top,
						fBtm = fRect.bottom;
					for (r = fTop; r <= fBtm; r++) {
						var srcRow = srcRows[r],
							dstRow = fRows[r];
						copyCells(left, fLeft, srcRow, dstRow);
					}
				}
				if (fRight < ar.rect.right) {
					//copy right cells
					var fRows = f.rows,
						fTop = fRect.top,
						fBtm = fRect.bottom;
					for (r = fTop; r <= fBtm; r++) {
						var srcRow = srcRows[r],
							dstRow = fRows[r];
						copyCells(fRect.right + 1, right, srcRow, dstRow);
					}
				}
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
	
	var ATTR_ALL = 1,
		ATTR_TEXT = 2,
		ATTR_STYLE = 3,
		ATTR_SIZE = 4,
		ATTR_MERGE = 5;
	function newCell(v, type) {
		var c = {
			/**
			 * Row number
			 */
			//r
			/**
			 * Column number
			 */
			//c
			/**
			 * Cell type
			 */
			//cellType,
			/**
			 * Cell text
			 */
			//text
			/**
			 * Cell edit text
			 */
			//editText,
			/**
			 * Cell format text
			 */
			//formatText
			/**
			 * Cell is locked or not
			 * 
			 * Default: true
			 */
			//lock
			/**
			 * whether the text should be wrapped or not
			 * 
			 * Default: false
			 */
			//wrap
			/**
			 * Horizontal alignment
			 * 
			 * <ul>
			 * 	<li>l: align left</li>
			 * 	<li>c: align center</li>
			 * 	<li>r: align right</li>
			 * </ul>
			 * 
			 * Default: "l"
			 */
			//halign
			/**
			 * Vertical alignment
			 * 
			 * <ul>
			 * 	<li>t: align top</li>
			 * 	<li>c: align center</li>
			 * 	<li>b: align bottom</li>
			 * </ul>
			 * 
			 * Default: "t"
			 */
			//valign
			/**
			 * Merge CSS class
			 */
			//mergeCls: v.mcls,
			/**
			 * Merge id
			 */
			//mergeId: v.mi,
			/**
			 * Merge rect
			 */
			//merge: null,
			/**
			 * Width id
			 */
			//widthId: v.w,
			/**
			 * Height id
			 */
			//heightId: v.h,
			/**
			 * Cell style
			 */
			//style
			/**
			 * Inner cell style
			 */
			//innerStyle
			/**
			 * Whether cell has right border or not
			 * 
			 * default: false
			 */
			//rightBorder
			update: function (v, type) {
				var upAll = type == ATTR_ALL,
					upText = (upAll || type == ATTR_TEXT),
					upStyle = (upAll || type == ATTR_STYLE),
					upSize = (upAll || type == ATTR_SIZE),
					upMerge = (upAll || type == ATTR_MERGE);
				if (upText) {
					var cellType = v.ct,
						mergedText = v.meft;
					this.cellType = cellType != undefined ? cellType : 3;
					if (mergedText != undefined) {
						this.text = this.editText = this.formatText = mergedText;
					} else {
						var text = v.t,
							editText = v.et
							formatText = v.ft;
						this.text = text != undefined ? text : '';
						this.editText = editText != undefined ? editText : '';
						this.formatText = formatText != undefined ? formatText : '';
					}
				}
				if (upStyle) {
					var style = v.s,
						innerStyle = v.is,
						wrap = v.wp,
						rbo = v.rb,
						lock = v.l,
						halign = v.ha,
						valign = v.va;
					this.style = style != undefined ? style : '';
					this.innerStyle = innerStyle != undefined ? innerStyle : '';
					this.wrap = wrap != undefined ? wrap == 't' : false;
					this.lock = lock != undefined ? lock : true;
					this.halign = halign != undefined ? halign : 'l';
					this.valign = valign != undefined ? valign : 't';
					this.rightBorder = rbo != undefined ? rbo == 't' : false;
				}
				if (upSize) {
					var wId = v.w,
						hId = v.h,
						drh = v.drh;
					this.widthId = wId;
					this.heightId = hId;
				}
				if (upMerge) {
					this.mergeId = v.mi;
					this.mergeCls = v.mcls;
					if (this.mergeId) {
						this.merge = newRect(v.mt, v.ml, v.mb, v.mr);
					}
				}
			}
		}
		c.update(v, type);
		return c;
	}
	
	function newRow(v, type, left, right) {
		var row = {
			r: v.r,
			heightId: v.h,
			cells: {},
			updateColumnWidthId: function (col, id) {
				var cell = this.cells[col];
				if (cell)
					cell.widthId = id;
			},
			updateRowHeightId: function (id) {
				var cells = this.cells;
				for (var p in cells) {
					cells[p].heightId = id;
				}
			},
			update: function (attr, type, left, right) {
				var src = attr.cs,
					i = left,
					j = 0,
					cell,
					r = this.r,
					cs = this.cells;
				for (; i <= right; i++) {
					var c = cs[i];
					if (!c) {
						c = cs[i] = newCell(src[j++], type);
						c.r = r;
						c.c = i;
					} else {
						c.update(src[j++], type);
					}
				}
			},
			removeColumns: function (col, size, rCol) {
				var cs = this.cells,
					i = size,
					lCol = col;
				for (var c = col; c <= rCol; c++) {
					var cell = cs[c];
					if (cell) {
						if (i > 0) {
							delete cs[c];
							i--;
						} else {
							delete cs[c];
							cell.c -= size; //re-index
							cs[cell.c] = cell;
						}
					}
				}
			},
			getCell: function (num) {
				return this.cells[num];
			}
		}
		row.update(v, type, left, right);
		return row;
	}
	
	function newRect(tRow, lCol, bRow, rCol) {
		return {
			top: tRow,
			left: lCol,
			bottom: bRow,
			right: rCol
		}
	}
	
	function updateHeaders(src, dest) {
		var headers = src.hs,
			i = src.s,
			end = src.e,
			j = 0;
		for (; i <= end; i++) {
			dest[i] = headers[j++];
		}
	}

	function newCachedRange(v) {
		var range = {
			topFrozen: null,
			leftFrozen: null,
			rows: {},
			rowHeaders: {},
			columnHeaders: {},
			//current range rect
			rect: null,
			updateColumnWidthId: function (col, id) {
				var r = this.rect,
					rows = this.rows,
					tRow = r.top,
					bRow = r.bottom;
				for (var r = tRow; r <= bRow; r++) {
					var row = rows[r];
					if (row)
						row.updateColumnWidthId(col, id);
				}
				if (this.leftFrozen) {
					this.leftFrozen.updateColumnWidthId(col, id);
				}
			},
			updateRowHeightId: function (row, id) {
				var row = this.rows[row];
				if (row)
					row.updateRowHeightId(id);
				if (this.topFrozen) {
					this.leftFrozen.updateRowHeightId(row, id);
				}
			},
			updateBoundary: function (dir, top, left, btm, right) {
				var rect = this.rect;
				if (!rect) {
					this.rect = newRect(top, left, btm, right);
					return;
				}
				else if (this.containsRange(top, left, btm, right)) {
					return;
				} else {
					var rect = this.rect;
					switch (dir) {
					case 'visible':
						rect.right = right;
						rect.bottom = btm;
						break;
					case 'jump':
						delete this.rect;
						delete this.rows;
						delete this.rowHeaders;
						delete this.columnHeaders;	
						
						this.rect = newRect(top, left, btm, right);
						this.rows = {};
						this.rowHeaders = {};
						this.columnHeaders = {};
						break;
					case 'east':
						rect.right = right;
						break;
					case 'west':
						rect.left = left;
						break;
					case 'south':
						rect.bottom = btm;
						break;
					case 'north':
						rect.top = top;
						break;
					}
				}
			},
			pruneLeft: function (size) {
				var rows = this.rows,
					left = this.rect.left,
					colHeaders = this.columnHeaders;
				for (var p in rows) {
					var r = rows[p],
						cs = r.cells,
						i = left,
						j = size;
					while (j--) {
						delete cs[i++];
					}
				}
				i = left;
				j = size;
				while (j--) {
					delete colHeaders[i++];
				}
				this.rect.left = left + size;
				if (this.topFrozen) {
					this.topFrozen.pruneLeft(size);
				}
			},
			pruneRight: function (size) {
				var rows = this.rows,
					right = this.rect.right,
					colHeaders = this.columnHeaders;
				for (var p in rows) {
					var r = rows[p],
						cs = r.cells,
						i = right,
						j = size;
					while (j--) {
						delete cs[i--];
					}
				}
				i = right,
				j = size;
				while (j--) {
					delete colHeaders[i--];
				}
				this.rect.right = right - size;
				if (this.topFrozen) {
					this.topFrozen.pruneRight(size);
				}
			},
			pruneTop: function (size) {
				var rows = this.rows,
					rowHeaders = this.rowHeaders,
					i = this.rect.top,
					j = size;
				while (j--) {
					delete rows[i];
					delete rowHeaders[i];
					i++;
				}
				if (this.leftFrozen) {
					this.leftFrozen.pruneTop(size);
				}
				this.rect.top += size;
			},
			pruneBottom: function (size) {
				var rows = this.rows,
					rowHeaders = this.rowHeaders,
					i = this.rect.bottom,
					j = size;
				while (j--) {
					delete rows[i];
					delete rowHeaders[i];
					i--;
				}
				if (this.leftFrozen) {
					this.leftFrozen.pruneBottom(size);
				}
				this.rect.bottom -= size;
			},
			containsRange: function (tRow, lCol, bRow, rCol) {
				var rect = this.rect;
				return	tRow >= rect.top && lCol >= rect.left &&
							bRow <= rect.bottom && rCol <= rect.right;
			},
			renameColumnHeaders: function (col, size, extnm) {
				var headers = this.columnHeaders,
					i = col,
					j = 0,
					l = extnm.length;
				while (j < l) {
					var h = headers[i++];
					if (h) {
						h.t = extnm[j];
					}
					j++;
				}
			},
			_updateHeaders: function (headers, index, size, extnm) {
				var	i = index,
					j = 0,
					l = extnm.length;
				while (j < l) {
					var h = headers[i++];
					if (h) {
						h.t = extnm[j];
					}
					j++;
				}
			},
			insertColumns: function (col, size, extnm) {
				this._updateHeaders(this.columnHeaders, col, size, extnm);
				this.rect.right += size;
			},
			removeColumns: function (col, size, extnm) {
				this._updateHeaders(this.columnHeaders, col, size, extnm);
				var r = this.rect,
					rows = this.rows,
					rCol = r.right,
					tRow = r.top,
					bRow = r.bottom;
				for (var r = tRow; r <= bRow; r++) {
					var row = rows[r];
					if (row) {
						row.removeColumns(col, size, rCol);
					}
				}
				this.rect.right -= size;
			},
			insertRows: function (row, size, extnm) {
				this._updateHeaders(this.rowHeaders, row, size, extnm);
				this.rect.bottom += size; 
			},
			removeRows: function (row, size, extnm) {
				this._updateHeaders(this.rowHeaders, row, size, extnm);
				var rows = this.rows,
					bRow = this.rect.bottom,
					i = size;
				for (var r = row; r <= bRow; r++) {
					var row = rows[r];
					if (row) {
						if (i > 0) {
							delete rows[r];
							i--;
						} else {
							delete rows[r];
							row.r -= size;
							rows[row.r] = row;
						}
					}
				}
				this.rect.bottom -= size;
			},
			update: function (v) {
				var attrType = v.at,
					top = v.t,
					left = v.l,
					btm = v.b,
					right = v.r,
					src = v.rs,
					rowHeaderObj = v.rhs,
					colHeaderObj = v.chs,
					i = top, 
					s = 0;
				
				this.updateBoundary(v.dir, top, left, btm, right);
				var rows = this.rows;
				for (; i <= btm; i++) {
					var row = rows[i];
					if (!row) {
						rows[i] = newRow(src[s++], attrType, left, right);
					} else {
						row.update(src[s++], attrType, left, right);
					}
				}
				
				if (rowHeaderObj) {
					updateHeaders(rowHeaderObj, this.rowHeaders);
				}
				if (colHeaderObj) {
					updateHeaders(colHeaderObj, this.columnHeaders);
				}
			},
			getRow: function (num) {
				return this.rows[num];
			}
		};
		range.update(v);
		return range;
	}
/**
 * Spreadsheet is a is a rich ZK Component to handle EXCEL like behavior
 */
var Spreadsheet =
zss.Spreadsheet = zk.$extends(zul.Widget, {
	/**
	 * Indicate Ctrl-Paste event key down status
	 */
	_ctrlPasteDown: false,
	$define: {
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
		preloadRowSize: null,
		preloadColSize: null,
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
	_activeRange: null,
	/**
	 * Sets active range
	 */
	setActiveRange: function (v) {
		var ar = this._activeRange;
		if (!ar) {
			this._activeRange = ar = newCachedRange(v);
			this.appendChild(this.sheetCtrl = new zss.SSheetCtrl());
		} else {
			var sheet = this.sheetCtrl,
				range;
			if (sheet) {
				ar.update(jq.evalJSON(v));
				this._triggerContentsChanged = true;
			}
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
				var sht = self.sheetCtrl;
				if (sht && sht._initiated)
					sht.dp.gainFocus(trigger);
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
		this._initMaxColumn();
		this._initMaxRow();
		this._initFrozenArea();

		var show = this._dpGridlines;
		if (typeof show != 'undefined')
			this.setDisplayGridlines(show);
	},
	bind_: function (desktop, skipper, after) {
		_calScrollWidth();
		this.$supers('bind_', arguments);
		this._initControl();
		zWatch.listen({onShow: this, onSize: this, onResponse: this});
	},
	unbind_: function () {
		zWatch.unlisten({onShow: this, onSize: this, onResponse: this});
		
		var r = this._activeRange;
		if (r) {
			this._activeRange = null;
		}
		this._maxColumnMap = this._maxRowMap = null;
		
		removeSSheet(this.uuid + "-sheet");
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
	CELL_MOUSE_EVENT_NAME: {lc:'onCellClick', rc:'onCellRightClick', dbc:'onCellDoubleClick', af:'onFilter', dv:'onValidateDrop'},
	HEADER_MOUSE_EVENT_NAME: {lc:'onHeaderClick', rc:'onHeaderRightClick', dbc:'onHeaderDoubleClick'},
	SRC_CMD_SET_COL_WIDTH: 'setColWidth',
	initLaterAfterCssReady: function (sheet) {
		if (zcss.findRule(".zs_indicator", sheet.sheetid + "-sheet") != null) {

			sheet._cmdGridlines(sheet._wgt.isDisplayGridlines());
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
				createSSheet("", sheet.sheetid + "-sheet-opera");//heigher
			
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