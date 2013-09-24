/* CacheCtrl.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 12, 2012 7:07:27 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
(function () {

	function newRect(tRow, lCol, bRow, rCol) {
		return {
			top: tRow,
			left: lCol,
			bottom: bRow,
			right: rCol
		}
	}
	
	function updateHeaders(dest, src) {
		var headers = src.hs,
			i = src.s,
			end = src.e,
			j = 0;
		for (; i <= end; i++) {
			var h = dest[i] = headers[j++];
			h.i = i;//header index
		}
	}
	
	var ATTR_ALL = 1,
		ATTR_TEXT = 2,
		ATTR_STYLE = 3,
		ATTR_SIZE = 4,
		ATTR_MERGE = 5;
	/**
	 * Create Cell model
	 * 
	 * @param JSON Object v data from server
	 * @param int type update type. Default ATTR_ALL means update all cell attributes
	 * @param int heightId height style id
	 * @param Array texts text string aggregation
	 * @param Array styles style string aggregation
	 * @param Object mergeInfos the merge info aggregation
	 */
	function newCell(v, type, heightId, texts, styles, mergeInfos) {
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
			 * Cell reference address 
			 */
			//ref
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
			 * Cell font size, default is 11pt 
			 */
			//fontSize
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
			/**
			 * Whether cell shall overflow or not
			 */
			//overflow
			/**
			 * Cell font size in point
			 * 
			 * default: 11
			 */
			//fontSize
			/**
			 * Update Cell model
			 * 
			 * @param JSON Object v data from server
			 * @param int type update type. Default ATTR_ALL means update all cell attributes
			 * @param int heightId height style id
			 * @param Array texts text string aggregation
			 * @param Array styles style string aggregation
			 * @param Object mergeInfos the merge info aggregation
			 */
			update: function (v, type, heightId, texts, styles, mergeInfos) {
				var upAll = type == ATTR_ALL,
					upText = (upAll || type == ATTR_TEXT),
					upStyle = (upAll || type == ATTR_STYLE),
					upSize = (upAll || type == ATTR_SIZE),
					upMerge = (upAll || type == ATTR_MERGE),
					cellType = v.ct;
				this.cellType = cellType != undefined ? cellType : 3;//default is BLANK_CELL
				if (upText) {
					var mergedTextId = v.meft;
					if (mergedTextId != undefined) {//index start from 0
						this.text = this.editText = this.formatText = texts[mergedTextId] || '';
					} else {
						var tId = v.t,
							eId = v.et
							fId = v.ft;
						this.text = tId != undefined ? texts[tId] : '';
						this.editText = eId != undefined ? texts[eId] : '';
						this.formatText = fId != undefined ? texts[fId] : '';
					}
				}
				if (upStyle) {
					var sId = v.s,
						isId = v.is,
						wrap = v.wp,
						rbo = v.rb,
						lock = v.l,
						halign = v.ha,
						valign = v.va,
						fontSize = v.fs;
					this.style = sId != undefined ? styles[sId] : '';
					this.innerStyle = isId != undefined ? styles[isId] : '';
					this.wrap = wrap != undefined;
					//bug ZSS-56: Unlock a cell, protect sheet, cannot double click to edit the cell
					this.lock = lock != undefined ? lock != 'f' : true;
					this.halign = halign || 'l'; //horizontal align: default left
					this.valign = valign || 'b'; //vertical align: default bottom 
					this.rightBorder = rbo != undefined;
					if (fontSize)
						this.fontSize = fontSize;
				}
				if (upSize) {
					// ZSS-224: modify overflow flag spec. to carry more status in bitswise format
					// refer to Spreadsheet.java -> getCellAttr()
					this.overflow = !!v.ovf; // backward compatibility (ovf flag is either undefined or greater than 0)
					this.overflowOpt = v.ovf ? v.ovf : 0; // if flag is undefined, the options flag should be zero
				}
				if (upMerge) {
					
					var idx = v.mi,
						cssIdx = v.mc;
					if (idx != undefined) {
						var info = mergeInfos.r[idx],
							css = mergeInfos.cs[cssIdx];
						this.mergeId = info.i;
						this.mergeCls = css;
						if (this.mergeId) {
							this.merge = newRect(info.t, info.l, info.b, info.r);
						}
					} else if (this.mergeId) {//remove merge
						delete this.mergeId;
						delete this.mergeCls;
						delete this.merge;
					}
				}
			}
		}
		c.update(v, type, heightId, texts, styles, mergeInfos);
		return c;
	}
	
	/**
	 * Create Row model
	 * 
	 * @param JSON Object v row data from server
	 * @param int update data type
	 * @param int update cell left number
	 * @param int update cell right number
	 * @param Array texts text string aggregation
	 * @param Array styles style string aggregation
	 * @param Object mergeInfos the merge info aggregation
	 */
	function newRow(v, type, left, right, texts, styles, mergeInfos) {
		var row = {
			r: v.r,
			heightId: v.h,
			cells: {},
			//TODO: rm this
			updateRowHeightId: function (id) {
				this.heightId = id;
				var cells = this.cells;
				for (var p in cells) {
					cells[p].heightId = id;
				}
			},
			update: function (attr, type, left, right, texts, styles, mergeInfos) {
				var src = attr.cs,
					i = left,
					j = 0,
					cell,
					r = this.r,
					cs = this.cells,
					hId = this.heightId;
				for (; i <= right; i++) {
					var c = cs[i];
					if (!c) {
						c = cs[i] = newCell(src[j++], type, hId, texts, styles, mergeInfos);
						c.r = r;
						c.c = i;
					} else {
						c.update(src[j++], type, hId, texts, styles, mergeInfos);
					}
					//row contains wrap cell may have height Id on client side
					if (!c.heightId && hId) {
						c.heightId = hId;
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
		row.update(v, type, left, right, texts, styles, mergeInfos);
		return row;
	}

/**
 * Cache data of sheet
 */
zss.ActiveRange = zk.$extends(zk.Object, {
	//id: sheet uuid
	/**
	 * Indicate shall prune data or not
	 */
	//prune: false
	/**
	 * Top frozen data
	 */
	//topFrozen
	/**
	 * Left frozen data
	 */
	//leftFrozen
	/**
	 * Current rectangle range
	 */
	rect: null,
	$init: function (data) {
		this.rows = {};
		this.rowHeaders = {};
		this.columnHeaders = {};
		this.update(data);
		this.id = data.id;
	},
	size: function () {
		var rect = this.rect,
			height = (rect.bottom - rect.top) + 1,
			width = (rect.right - rect.left) + 1;
		return height * width;
	},
	updateColumnWidthId: function (col, id) {
		var r = this.rect,
			tRow = r.top,
			bRow = r.bottom,
			rows = this.rows,
			header = this.columnHeaders[col];
		if (header)
			header.p = id;

		if (this.leftFrozen) {
			this.leftFrozen.updateColumnWidthId(col, id);
		}
	},
	getColumnWidthId: function (col) {
		var c = this.columnHeaders[col];
		return c ? c.p : null;
	},
	updateRowHeightId: function (row, id) {
		var r = this.rows[row],
			header = this.rowHeaders[row];
		if (r)
			r.updateRowHeightId(id);
		if (header)
			header.p = id;
		if (this.topFrozen) {
			this.topFrozen.updateRowHeightId(row, id);
		}
	},
	getRowHeightId: function (row) {
		var r = this.rowHeaders[row];
		return r ? r.p : null; 
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
				//row contains wrap cell may have height Id on client side, delete it later
				//delete this.rows;
				//delete this.rowHeaders;
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
	insertNewColumn: function (colIdx, size, headers) {
		updateHeaders(this.columnHeaders, headers);
		var rows = this.rows;
			rng = this.rect,
			rCol = rng.right;
		for (var r = rng.top; r <= rng.bottom; r++) {
			var cs = rows[r].cells,
				c = colIdx,
				cb = colIdx + size,//column index boundary 
				ccs = [],//clone cells
				cfn = zss.ActiveRange.clone;
			//clone cells for insert later (after shift)
			for (; c < cb; c++) {
				ccs.push(cfn(cs[c]));
			}
			
			//shift cells right
			cb = colIdx - 1;
			c = rCol;
			while (c != cb) {
				var cell = cs[c--],
					oIdx = cell.c,//old index
					nIdx = oIdx + size;//new index
				cell.c = nIdx;
				cs[nIdx] = cell;
				cs[oIdx] = null;
			}
			
			//insert cells form clone
			var cc;
			c = colIdx;
			while (cc = ccs.shift()) {
				cs[c++] = cc; 
			}
		}
		this.rect.right += size;
	},
	removeColumns: function (col, size, headers) {
		updateHeaders(this.columnHeaders, headers);
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
	insertNewRow: function (rowIdx, size, headers) {
		updateHeaders(this.rowHeaders, headers);
		
		var rows = this.rows;
			rng = this.rect,
			lCol = rng.left,
			rCol = rng.right,
			r = rowIdx,
			rb = rowIdx + size,
			crs = [], //clone rows
			cfn = zss.ActiveRange.copyRow;

		//cloned rows for insert later (after shift row)
		for (;r < rb; r++) {
			crs.push(cfn(lCol, rCol, rows[r]));
		}
		
		//shift rows
		rb = rowIdx - 1;
		r = rng.bottom;
		while (r != rb) {
			var row = rows[r--],
				oIdx = row.r,//old index
				nIdx = oIdx + size;//new index
			row.r = nIdx;
			rows[nIdx] = row;
			rows[oIdx] = null;
		}
		
		//insert rows from clone
		var ro;
		r = rowIdx;
		while (ro = crs.shift()) {
			rows[r++] = ro;
		}
			
		this.rect.bottom += size; 
	},
	removeRows: function (row, size, headers) {
		updateHeaders(this.rowHeaders, headers);
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
	//TODO: update overlap range
	fetchUpdate: function (v) {
		var top = v.t,
			left = v.l,
			btm = v.b,
			right = v.r,
			rect = this.rect;
		if (rect.top == btm + 1) { //update north
			v.dir = 'north';
			this.update(v);
		} else if (rect.left == right + 1) {//update left
			v.dir = 'west';
			this.update(v);
		} else if (rect.right == left - 1) {//update right
			v.dir = 'east';
			this.update(v);
		} else if (rect.bottom == top - 1) {//update south
			v.dir = 'south';
			this.update(v);
		}
	},
	update: function (v, dir) {

		// ZSS-392: multiple ActiveRange for different panels
		// just delegate to corresponding freeze active ranges 
		if(v.leftFrozen) { // left frozen data
			if(this.leftFrozen) {
				this.leftFrozen.update(v.leftFrozen, dir);
			} else {
				this.leftFrozen = new zss.FreezeActiveRange(v.leftFrozen);
			}
		}
		if(v.topFrozen) { // top frozen data
			if(this.topFrozen) {
				this.topFrozen.update(v.topFrozen, dir);
			} else {
				this.topFrozen = new zss.FreezeActiveRange(v.topFrozen);
			}
		}
		if(v.cornerFrozen) { // corner frozen data
			if(this.cornerFrozen) {
				this.cornerFrozen.update(v.cornerFrozen, dir);
			} else {
				this.cornerFrozen = new zss.FreezeActiveRange(v.cornerFrozen);
			}
		}
		
		// real update cells
		this.updateCells(v, dir);
	},
	updateCells: function (v, dir) {
		var attrType = v.at;
		var top = v.t;
		var left = v.l;
		var btm = v.b;
		var right = v.r;
		var src = v.rs;
		var textAggregation = v.s;
		var styleAggregation = v.st;
		var mergeAggregation = v.m;
		var rowHeaderObj = v.rhs;
		var colHeaderObj = v.chs;
		var i = top; 
		var s = 0;
		var dir = dir || v.dir;
		var oldRow;
		var oldRows = {};
		var oldRowHeaders = {};
		
		if ('jump' == dir) {
			//row contains wrap cell may have height Id on client side
			oldRows = this.oldRows = this.rows,
			oldRowHeaders = this.oldRowHeaders = this.rowHeaders;
		}
		this.updateBoundary(dir, top, left, btm, right);
		
		var rows = this.rows;
		for (; i <= btm; i++) {
			var row = rows[i];
			if (!row) {
				row = rows[i] = newRow(src[s++], attrType, left, right, textAggregation, styleAggregation, mergeAggregation);
				//row contains wrap cell may have height Id on client side
				if ('jump' == dir) {
					oldRow = oldRows[i];
					if (oldRow && oldRow.heightId && !row.heightId) {
						row.updateRowHeightId(oldRow.heightId);
					}
				}
			} else {
				row.update(src[s++], attrType, left, right, textAggregation, styleAggregation, mergeAggregation);
			}
		}
		
		if (rowHeaderObj) {
			updateHeaders(this.rowHeaders, rowHeaderObj);
			//row contains wrap cell may have height Id on client side
			if ('jump' == dir) {
				var headers = this.rowHeaders;
				for (var i in headers) {
					var h = headers[i],
						oldHeader = oldRowHeaders[i];
					if (!h.p && oldHeader && oldHeader.p) {
						h.p = oldHeader.p; //position id
					}
				}
			}
		}
		
		if (colHeaderObj) {
			updateHeaders(this.columnHeaders, colHeaderObj);
		}
		
		//row contains wrap cell may have height Id on client side
		if ('jump' == dir) {
			delete this.oldRows;
			delete this.oldRowHeaders;
		}
	},
	getRow: function (num) {
		return this.rows[num];
	},
	clone: function (tRow, lCol, bRow, rCol, src) {
		var rows = this.rows,
			rhs = this.rowHeaders,
			chs = this.columnHeaders,
			srcRowHeaders = src.rowHeaders,
			srcColHeaders = src.columnHeaders,
			cpRowsFn = zss.ActiveRange.copyRow,
			cfn = zss.ActiveRange.clone;
		for (var r = tRow; r <= bRow; r++) {
			var sRow = src.getRow(r);
			rows[r] = cpRowsFn(lCol, rCol, sRow);
			rhs[r] = cfn(srcRowHeaders[r]);//clone row headers
		}
		
		for (var c = lCol; c <= rCol; c++) {
			chs[c] = cfn(srcColHeaders[c]);
		}
	}
}, {//static
	copyRow: function (lCol, rCol, srcRow) {
		var row = {
			r: srcRow.r,
			heightId: srcRow.heightId,
			cells: {},
			update: srcRow.update,
			getCell: srcRow.getCell,
			updateRowHeightId: srcRow.updateRowHeightId,
			removeColumns: srcRow.removeColumns
		};
		zss.ActiveRange.copyCells(lCol, rCol, srcRow, row);
		return row;
	},
	copyCells: function (lCol, rCol, srcRow, dstRow) {
		var srcCells = srcRow.cells,
			dstCells = dstRow.cells,
			fn = zss.ActiveRange.clone;
		for (var c = lCol; c <= rCol; c++) {
			dstCells[c] = fn(srcCells[c]);
		}
	},
	clone: function (src) {
		var c = {};
		for (var p in src) {
			c[p] = src[p];
		}
		return c;
	}
});

	function copyAttributes(dst, src, fields) {
		for (var key in fields) {
			var f = fields[key];
			dst[f] = src[f];
		}
	}
/**
 * Snapshot sheet relate status and attributes
 * 
 * <ul>
 * 	<li>sheet style</li>
 * 	<li>row freeze</li>
 * 	<li>column freeze</li>
 * 	<li>row height</li>
 * 	<li>column width</li>
 * 	<li>visible range</li>
 *  <li>focus</li>
 * 	<li>selection</li>
 *  <li>highlight</li>
 * 	<li>displayGridlines</li>
 * 	<li>protect</li>
 * 	<li></li>
 * </ul>
 */
zss.Snapshot = zk.$extends(zk.Object, {
	$init: function (wgt) {
		var sheet = wgt.sheetCtrl,
			dataPanel = sheet.dp,
			leftPanel = sheet.lp,
			topPanel = sheet.tp,
			scrollPanel = sheet.sp,
			visRng = zss.SSheetCtrl._getVisibleRange(sheet);
		copyAttributes(this, wgt, 
			['_scss', '_displayGridlines', '_rowFreeze', '_columnFreeze', '_rowHeight', '_columnWidth', '_protect']);
		
		this.setCustRowHeight(sheet.custRowHeight.custom);
		this.setCustRowLastId(sheet.custRowHeight.ids.last);
		this.setCustColWidth(sheet.custColWidth.custom);
		this.setCustColLastId(sheet.custColWidth.ids.last);
		this.setMergeMatrix(sheet.mergeMatrix.mergeMatrix);
		this.setVisibleRange(visRng);
		this.setFocus(sheet.getLastFocus());
		this.setSelection(sheet.getLastSelection());
		if (sheet.isHighlightVisible()) {
			this.setHighlight(sheet.getLastHighlight());
		}
		
		this.setDataPanelSize({'width': dataPanel.width, 'height': dataPanel.height});
		this.setScrollPanelPos({'scrollLeft': scrollPanel.currentLeft, 'scrollTop': scrollPanel.currentTop});
		this.setLeftPanelPos(leftPanel.toppos);
		this.setTopPanelPos(topPanel.leftpos);
		
		if (wgt.getDataValidations) {
			var dv = wgt.getDataValidations();
			if (dv) {
				this.setDataValidations(dv);
			}
		}
		if (wgt.getAutoFilter) {
			var af = wgt.getAutoFilter();
			if (af) {
				this.setAutoFilter(af);
			}
		}
	},
	$define: {
		scss: null,
		rowFreeze: null,
		columnFreeze: null,
		rowHeight: null,
		columnWidth: null,
		protect: null,
		displayGridlines: null,
		/**
		 * @param array
		 */
		custRowHeight: null,
		/**
		 * @param int
		 */
		custRowLastId:null,
		/**
		 * @param array
		 */
		custColWidth: null,
		/**
		 * @param int
		 */
		custColLastId:null,
		/**
		 * @param array
		 */
		mergeMatrix: null,
		visibleRange: null,
		/**
		 * Data panel's width/height
		 */
		dataPanelSize: null,
		/**
		 * Scroll panel's scroll left and scroll top position
		 */
		scrollPanelPos: null,
		/**
		 * Left panel's top position
		 */
		leftPanelPos: null,
		/**
		 * Top panel's left position
		 */
		topPanelPos: null,
		focus: null,
		selection: null,
		highlight: null,
		dataValidations: null,
		autoFilter: null
	}
});


/**
 * Control cache data of sheets
 */
zss.CacheCtrl = zk.$extends(zk.Object, {
	/**
	 * Current sheet data
	 */
	selected: null,
	$init: function (wgt, v) {
		this._wgt = wgt;
		this.sheet = wgt.sheetCtrl;
		
		//key: sheet uuid, value: zss.ActiveRange
		this.sheets = {};
		
		//key: sheet uuid, value: sheet last status
		this.snapshots = {};
		
		this.setSelectedSheet(v);
	},
	/**
	 * Save current sheet status
	 */
	snap: function (sheetId) {
		this.snapshots[sheetId] = new zss.Snapshot(this._wgt);
	},
	getSnapshot: function (sheetId) {
		return this.snapshots[sheetId];
	},
	isCached: function (sheetId) {
		return !!this.sheets[sheetId];
	},
	setSelectedSheetBy: function (sheetId) {
		this.selected = this.sheets[sheetId];
	},
	getSheetBy: function (shtId) {
		return this.sheets[shtId];
	},
	setSelectedSheet: function (v) {
		var sheetId = v.id,
			rng = this.sheets[sheetId] = new zss.ActiveRange(v);
		
		this.selected = rng;
	},
	getSelectedSheet: function () {
		return this.selected;
	}
});

//ZSS-392: update freeze panels' activeRange individually
zss.FreezeActiveRange = zk.$extends(zss.ActiveRange, {

	$init: function (data) {
		this.$supers(zss.FreezeActiveRange, '$init', [data]); 
	},
	
	// override
	update: function (v, dir) {
		// just update cells
		this.updateCells(v, dir);
	}
	
}); // end of zss.FreezeActiveRange

})();