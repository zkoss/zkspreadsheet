/* MainBlockCtrl.js

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
	var IDLE = 0,
		LOADING = 1;
/**
 * MainBlockCtrl handle scroll event and load spreadsheet content
 */
zss.MainBlockCtrl = zk.$extends(zss.CellBlockCtrl, {
	loadstate: IDLE,
	replaceWidget: function (newwgt, leftFrozen, topFrozen) {
		this.$supers(zss.MainBlockCtrl, 'replaceWidget', [newwgt]); //create cells
		
		var r = newwgt.range;
		newwgt.create_('jump', r.top, r.left, r.bottom, r.right, leftFrozen, topFrozen, true);//wgt init won't create frozen area
	},
	_createVer: function (dir, lCol, rCol, topFrozen) {
		var tp = this.sheet.tp,
			data = topFrozen ? topFrozen.data : null,
			frozenRowStart = data ? data.t : -1,
			frozenRowEnd = data ? data.b : -1;
		tp.create_(dir, lCol, rCol, frozenRowStart, frozenRowEnd);
	},
	_createHor: function (dir, tRow, bRow, leftFrozen) {
		var lp = this.sheet.lp,
			data = leftFrozen ? leftFrozen.data : null,
			frozenColStart = data ? data.l : -1,
			frozenColEnd = data ? data.r : -1;
		lp.create_(dir, tRow, bRow, frozenColStart, frozenColEnd);
	},
	/**
	 * Create cells and associated headers
	 */
	create_: function (dir, tRow, lCol, bRow, rCol, leftFrozen, topFrozen, createFrozenOnly) {
		var fixSize = false;
		switch (dir) {
		case 'east':
			fixSize = true;
		case 'west':
			this._createVer(dir, lCol, rCol, topFrozen);
			break;
		case 'north':
			fixSize = true;
		case 'south':
			this._createHor(dir, tRow, bRow, leftFrozen);
			break;
		case 'jump':
			this._createVer(dir, lCol, rCol, topFrozen);
			this._createHor(dir, tRow, bRow, leftFrozen);
			break;
		}
		if (!createFrozenOnly)
			this.$supers(zss.MainBlockCtrl, 'create_', [dir, tRow, lCol, bRow, rCol]); //create cells;
		
		if (fixSize)
			this.sheet.dp._fixSize(this);
	},
	/**
	 * Create cells from cache
	 * 
	 * @param string dir direction
	 * @param int prune reserve when prune, -1 means don't prune, 0 means don't reserve 
	 */
	_createCellsIfCached: function (dir, size, jump) {
		var sheet = this.sheet,
			cr = this.range,
			ar = sheet._wgt._activeRange,
			vr = zss.SSheetCtrl._getVisibleRange(sheet);
		switch (dir) {
		case 'south':
			var tRow = cr.bottom + 1,
				lCol = cr.left,
				rCol = cr.right,
				bRow = tRow + size - 1;
			bRow = Math.min(bRow, sheet.maxRows - 1);
			if (ar.containsRange(tRow, lCol, bRow, rCol)) {
				this.create_(dir, tRow, lCol, bRow, rCol);
				if (cr.top + vr.height < vr.top) {
					this.pruneCell('north', vr, jump ? null : vr.top - (cr.top + vr.height));
				}
				return true; 
			}
			break;
		case 'north':
			var bRow = cr.top - 1,
				lCol = cr.left,
				rCol = cr.right,
				tRow = bRow - size + 1;
			bRow = bRow >= 0 ? bRow : 0;
			tRow = tRow >= 0 ? tRow : 0;
			if (ar.containsRange(tRow, lCol, bRow, rCol)) {
				this.create_(dir, tRow, lCol, bRow, rCol);
				if (cr.bottom - vr.height > vr.bottom) {
					this.pruneCell('south', vr, jump ? null : cr.bottom - vr.height - vr.bottom);
				}
				return true; 
			}
			break;
		case 'west':
			var tRow = cr.top,
				bRow = cr.bottom,
				rCol = cr.left - 1,
				lCol = rCol - size + 1;
			rCol = rCol >= 0 ? rCol : 0;
			lCol = lCol >= 0 ? lCol : 0;
			if (ar.containsRange(tRow, lCol, bRow, rCol)) {
				this.create_(dir, tRow, lCol, bRow, rCol);
				if (cr.right - vr.width > vr.right) {
					this.pruneCell('east', vr, jump ? null : cr.right - vr.width - vr.right);
				}
				return true; 
			}	
			break;
		case 'east':
			var tRow = cr.top,
				bRow = cr.bottom,
				lCol = cr.right + 1,
				rCol = lCol + size - 1;
			rCol = Math.min(rCol, sheet.maxCols - 1);
			if (ar.containsRange(tRow, lCol, bRow, rCol)) {
				this.create_(dir, tRow, lCol, bRow, rCol)
				if (cr.left + vr.width < vr.left) {
					this.pruneCell('west', vr, jump ? null : (vr.left - (cr.left + vr.width)));
				}
				return true; 
			}
			break;
		}
		return false;
	},
	/**
	 * Handle scroll on spreadsheet.
	 */
	doScroll: function (vertical) {
		var sheet = this.sheet,
			range = zss.SSheetCtrl._getVisibleRange(sheet),
			alwaysjump = false;
		/*if(!sheet.snapedBlock && sheet.snapedBlock!=sheet.activeBlock){
			sheet._snapActiveBlock();
			alwaysjump = true; 
		}*/
		if (vertical) {
			var ctop = this.range.top,
				cbottom = this.range.bottom;
			if (range.top >= ctop && range.top <= cbottom) {
				if(range.bottom < cbottom)
					return; //the visible is be contained.
				//neighbor south
				if (!this._createCellsIfCached('south', range.height)) {
					sheet.activeBlock.loadCell(range.bottom, range.left, (alwaysjump ? -1 : 20), null, alwaysjump);
				}
			} else if(ctop >= range.top && ctop <= range.bottom) {
				//neighbor north;
				if (!this._createCellsIfCached('north', range.height)) {
					sheet.activeBlock.loadCell(range.top,range.left, (alwaysjump ? -1 : 20), null, alwaysjump);
				}
			} else if(range.top > cbottom) {
				//jump south
				if (!this._createCellsIfCached('south', range.height, true)) {
					sheet.activeBlock.loadCell(range.bottom, range.left,-1, null, alwaysjump);
				}
			} else if(ctop > range.bottom) {
				//jump north;
				if (!this._createCellsIfCached('north', range.height, true)) {
					sheet.activeBlock.loadCell(range.top, range.left, -1, null, alwaysjump);
				}
			} else{
				return;
			}
		} else {
			var cleft = this.range.left,
				cright = this.range.right;
			if (range.left >= cleft && range.left <= cright) {
				if (range.bottom < cbottom)
					return; //the visible is be contained.
				//neighbor east
				if (!this._createCellsIfCached('east', range.width)) {
					sheet.activeBlock.loadCell(range.top, range.right, (alwaysjump ? -1 : 5), null, alwaysjump);
				}
			} else if (cleft >= range.left && cleft <= range.right) {
				//neighbor west;
				if (!this._createCellsIfCached('west', range.width)) {
					sheet.activeBlock.loadCell(range.top, range.left, (alwaysjump ? -1 : 5), null, alwaysjump);
				}
			} else if (range.left>cright) {
				//jump east
				if (!this._createCellsIfCached('east', range.width, true)) {
					sheet.activeBlock.loadCell(range.top, range.right, -1, null, alwaysjump);
				}
			} else if (cleft > range.right) {
				//jump west;
				if (!this._createCellsIfCached('west', range.width, true)) {
					sheet.activeBlock.loadCell(range.top, range.left, -1, null, alwaysjump);
				}
			} else {
				return;
			}
		}
	},
	_sendOnCellFetch: function(token, type, direction, fetchLeft, fetchTop, fetchWidth, fetchHeight, vrange) {
		var sheet = this.sheet;
		this.loadstate = zss.MainBlockCtrl.LOADING;
		if (!vrange)
			vrange = zss.SSheetCtrl._getVisibleRange(sheet);

		var range = this.range,
			spcmp = this.sheet.sp.comp,
			dp = this.sheet.dp,
			arTopHgh = -1,
			arBtmHgh = -1
			arLeftWidth = -1,
			arRightWidth = -1;
		if (('east' == direction || 'west' == direction) && this.sheet._wgt._preloadRowSize > 0) {
			var	ar = sheet._wgt._activeRange,
				topHgh = range.top - ar.rect.top + 1,
				btmHgh = ar.rect.bottom - range.bottom + 1;
			if (topHgh > 1)
				arTopHgh = topHgh;
			if (btmHgh > 1)
				arBtmHgh = btmHgh;
		}
		//TODO: if the amount of cell to load is small, shall load extra cells, not prune cached cells.
		if ('north' == direction || 'south' == direction) {
			var	ar = sheet._wgt._activeRange,
				leftWidth = range.left - ar.rect.left,
				rightWidth = ar.rect.right - range.right;
			
			if (leftWidth > 0) {
				ar.pruneLeft(leftWidth);
			}
			if (rightWidth > 0) {
				ar.pruneRight(rightWidth);
			}
		}
			
		sheet._wgt.fire('onZSSCellFetch', 
		 {token: token, sheetId: sheet.serverSheetId, type: type, direction: direction,
		 dpWidth: dp.width, dpHeight: dp.height, viewWidth: spcmp.clientWidth, viewHeight: spcmp.clientHeight,
		 blockLeft: range.left, blockTop: range.top, blockRight: range.right, blockBottom: range.bottom,
		 fetchLeft: fetchLeft, fetchTop: fetchTop, fetchWidth: fetchWidth, fetchHeight: fetchHeight,
		 rangeLeft: vrange.left, rangeTop: vrange.top, rangeRight: vrange.right, rangeBottom: vrange.bottom, 
		 arFetchTopHeight: arTopHgh, arFetchBtmHeight: arBtmHgh}, {toServer: true}, 25);
	},
	/**
	 * Returns Load Cell Direction
	 * @return string last direction (H: horizon; V: vertical; or null)
	 */
	_getLastDirection: function () {
		return this.sheet.activeBlock._lastDir;
	},
	/**
	 * load cell , if cell of row/col not exist, then load the cell into this block
	 * @param {Object} row row of a cell
	 * @param {Object} col col of a cell
	 * @param {Object} pruneres reserve when prune, -1 means don't prune, 0 means don't reserve 
	 * @param {Object} callbackfn callback function after cell be loaded
	 * @param {Object} alwaysjump always jump to this cell, this means skip this block, use a new bolck instead.
	 * @return true if already loaded, false if need to asynchronize loading.
	 */
	loadCell: function (row, col, pruneres, callbackfn, alwaysjump) {
		var cleft = this.range.left,
			ctop = this.range.top,
			cw = this.range.width,
			ch = this.range.height,
			cright = this.range.right,//cleft + cw - 1;
			cbottom = this.range.bottom;//ctop + ch - 1;
		
		if (row >= ctop && row <= cbottom && col >= cleft && col <= cright)
			return true;
		if (this.loadstate != zss.MainBlockCtrl.IDLE)//still waiting previous loading.
			return false;
		var token = "";
			sheet = this.sheet,
			range = zss.SSheetCtrl._getVisibleRange(sheet),
			local = this,
			fetchw = range.width, //size of cell of width to fetch back
			fetchh = range.height; //size of cell of height to fetch back
		
		if ((row >= ctop && row <= cbottom)) {//horizontal shift, east or west
			if (col < cleft) {//minus, west
				var y = ctop,
					h = ch,
					x = cleft - 1;
					w = x - col + 1;
				
				if (alwaysjump || w > fetchw) {
					//jump;
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "west", col, row, -1, -1, range);
				} else {
					//neighbor
					w = fetchw;
					if (x - w < 0)
						w = x + 1;

					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock,//get current activeBlock again.
							lastDir = block._lastDir; 
						block._lastDir = "H";
						if (lastDir && lastDir == "V"){
							block.pruneCell("south", range, 1);
							block.pruneCell("north", range, 1);
						} else if(pruneres >= 0) {
							block.pruneCell("east", range, pruneres);
						}
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if(callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "west", x, y, w, h, range);
				}

			} else {//positive, east
				var y = ctop,
					x = cleft + cw,
					w = col-x + 1;
				if (alwaysjump || w > fetchw) {
					//jump;
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "east", col, row, -1, -1, range);
				} else {
					//neighbor
					w = fetchw;
					var h = ch;
					if (x + w > this.sheet.maxCols)
						w = this.sheet.maxCols - x;
	
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock,//get current activeBlock again.
							lastDir = block._lastDir; 
						block._lastDir = "H";
						if (lastDir && lastDir == "V") {
							block.pruneCell("south", range, 1);
							block.pruneCell("north", range, 1);
						} else if (pruneres >= 0){
							block.pruneCell("west", range, pruneres);
						}
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "east", x, y, w, h, range);
				}
			}
		} else if ((col >= cleft && col <=cright)) {//vertical shift, shouth or north
			if (row < ctop) {//minus, north
				var x = cleft,
					w = cw,
					y = ctop - 1,
					h = y - row + 1;
				if (alwaysjump || h > fetchh) {
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "north", col, row, -1, -1, range);
				} else {
					h = fetchh;
					if (y - h < 0)
						h = y + 1;

					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock,//get current activeBlock again.
							lastDir = block._lastDir; 
						local._lastDir = "V";
						if (lastDir && lastDir == "H") {
							block.pruneCell("east", range, 1);
							block.pruneCell("west", range, 1);
						} else if (pruneres >= 0) {
							block.pruneCell("south", range, pruneres);
						}
						sheet.activeBlock.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();

						sheet.activeBlock.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "north", x, y, w, h, range);
				}
			} else {//positive, south
				var y = ctop + ch,
					x = cleft,
					w = cw,
					h = row-y +1 ;

				if (alwaysjump || h > fetchh) {
					//jump
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "south", col, row, -1, -1, range);
				} else {
					h = fetchh;
					if (y + h > this.sheet.maxRows)
						h = this.sheet.maxRows - y;
					
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						var lastDir = block._lastDir; 
						block._lastDir = "V";
						if (lastDir && lastDir == "H") {
							block.pruneCell("east", range, 1);
							block.pruneCell("west", range, 1);
						} else if(pruneres >= 0) {
							block.pruneCell("north", range, pruneres);
						}
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "south", x, y, w, h, range);
				}
			}
		} else {
			token = zkS.addCallback(function () {
				if(sheet.invalid) return;
				var block = sheet.activeBlock ;//get current activeBlock again.
				block.loadstate = zss.MainBlockCtrl.IDLE;
				if(callbackfn) callbackfn();
				//local variable will be destoryed, don't use local.loadForVisible
				block.loadForVisible();
			});

			var direction = "";
			if (row < ctop && col < cleft) {
				//west-north
				direction = "westnorth";
			} else if(row < ctop && col > cright) {
				//east-north
				direction = "eastnorth";
			} else if(row > cbottom && col < cleft) {
				//west-south
				direction = "westsouth";
			} else if(row > cbottom && col > cright) {
				//east-shouth
				direction = "eastsouth";
			} else{
				direction = "eastnorth";
			}
			this._sendOnCellFetch(token, "jump", direction, col, row, -1, -1, range);
		}
		return false;
		
	},
	/**
	 * Prune cell
	 * @param string type the direction to prune
	 * @param zss.Range range the range to prune
	 * @param int reserve
	 */
	pruneCell: function (type , range, reserve) {
		var sheet = this.sheet;
		if (!sheet.config.prune) return;
		
		if (!range)
			range = zss.SSheetCtrl._getVisibleRange(sheet);
		if (!reserve || reserve < 0)
			reserve = 0;
		
		var sync = false;
		var wgt = sheet._wgt,
			ar = wgt._activeRange,
			rect = ar.rect;
		if (type == "west") {
			var l = range.left - reserve;
			l = sheet.mergeMatrix.getLeftConnectedColumn(l, this.range.top, this.range.bottom);

			if (l <= 0 || l < this.range.left) return;
			var size = l - this.range.left;
			if (size > 0) {
				this.removeColumnsFromStart_(size);
				sheet.tp.removeChildFromStart_(size);
				var colSize = wgt._preloadColSize;
				if (colSize <= 0)
					ar.pruneLeft(size);
				else {
					var pruneSize = (l - ar.rect.left + 1) - colSize;
					if (pruneSize > 0)
						ar.pruneLeft(pruneSize);
				}
				sync = true;
			}
		} else if (type == "east") {
			var r = range.right + reserve;
			r = sheet.mergeMatrix.getRightConnectedColumn(r,this.range.top,this.range.bottom);
			
			if(r > this.range.right) return;
			var size = this.range.right - r;
			if (size > 0) {
				this.removeColumnsFromEnd_(size);
				sheet.tp.removeChildFromEnd_(size);
				var colSize = wgt._preloadColSize;
				if (colSize <= 0)
					ar.pruneRight(size);
				else {
					var pruneSize = (ar.rect.right - r + 1) - colSize;
					if (pruneSize > 0)
						ar.pruneRight(pruneSize);
				}
				sync = true;
			}
		} else if (type == "north") {
			var t = range.top - reserve;
			t = sheet.mergeMatrix.getTopConnectedRow(t, this.range.left, this.range.right);
			
			if(t <= 0 || t < this.range.top) return;
			var size = t - this.range.top;
			if (size > 0) {
				this.removeRowsFromStart_(size);
				sheet.lp.removeChildFromStart_(size);
				var rowSize = wgt._preloadRowSize;
				if (rowSize <= 0) {
					ar.pruneTop(size);
				} else {
					var pruneSize = (t - ar.rect.top + 1) - rowSize;
					if (pruneSize > 0)
						ar.pruneTop(pruneSize);
				}
				sync = true;
			}
		} else if (type == "south") {
			var b = range.bottom + reserve;
			b = sheet.mergeMatrix.getBottomConnectedRow(b,this.range.left,this.range.right);
			
			if (b > this.range.bottom) return;
			var size = this.range.bottom - b;
			if (size > 0) {
				this.removeRowsFromEnd_(size);
				sheet.lp.removeChildFromEnd_(size);
				var rowSize = wgt._preloadRowSize;
				if (rowSize <= 0) {
					ar.pruneTop(size);
				} else {
					var pruneSize = (ar.rect.bottom - b + 1) - rowSize;
					if (pruneSize > 0)
						ar.pruneBottom(pruneSize);
				}
			}
			sync = true;
		}
		if (sync) {
			//bug 1951423 IE : row is broken when scroll down
			//i must use timeout to delay processing
			setTimeout(function () {
				sheet.sendSyncblock();
			}, 0);	
		}
		sheet.dp._fixSize(this);
	},
	_getFetchWidth: function (rCol) {
		var sheet = this.sheet;
		return Math.min(sheet._wgt.getMaxColumn(), rCol + Math.round(sheet._wgt._preloadColSize / 2)) - this.range.right;
	},
	_getFetchHeight: function (bRow) {
		var sheet = this.sheet;
		return Math.min(sheet._wgt.getMaxRow(), bRow + Math.round(sheet._wgt._preloadRowSize / 2)) - this.range.bottom;
	},
	/**
	 * Load content and set the range to visible 
	 * @param zss.Range 
	 */
	loadForVisible: function (vrange) {
		var local = this,
			sheet = this.sheet;
		if (!vrange)
			vrange = zss.SSheetCtrl._getVisibleRange(sheet);
		
		//Two phases
		//1. create cells from cache if possible
		//2. fetch data from server
		
		var tRow = vrange.top,
			lCol = vrange.left,
			rCol = vrange.right,
			bRow = vrange.bottom,
			range = this.range,
			top = range.top,
			left = range.left,
			right = range.right,
			bottom = range.bottom,
			ar = sheet._wgt._activeRange,
			created = false;
		
		//create east from cache
		if (right + 1 <= rCol) {
			var chd = false;
			if (ar.containsRange(top, right + 1, bottom, rCol)) {
				this.create_('east', top, right + 1, bottom, rCol);
				chd = true;
			} else if (ar.rect.right > right + 1 && ar.containsRange(top, right + 1, bottom, ar.rect.right)) {	
				//create partial east from cache
				this.create_('east', top, right + 1, bottom, ar.rect.right);
				chd = true;
			}
			if (chd) { //after create cell from cache, range's value may changed
				created = true;
				range = this.range;
				top = range.top;
				left = range.left;
				right = range.right;
				bottom = range.bottom;
			}
		}
		//create south from cache
		if (bottom + 1 <= bRow) {
			var chd = false;
			if (ar.containsRange(bottom + 1, lCol, bRow, rCol)) {
				this.create_('south', bottom + 1, lCol, bRow, rCol);
				chd = true;
			} else if (ar.containsRange(bottom + 1, lCol, bRow, right)) {
				//create partial south; need to fetch south-east range
				this.create_('south', bottom + 1, lCol, bRow, right);
				chd = true;
			}
			if (chd) {
				created = true;
				range = this.range;
				top = range.top;
				left = range.left;
				right = range.right;
				bottom = range.bottom;
			}
		}
		
		if (tRow > bottom || bRow < top || lCol > right || rCol < left) {
			//visible range not cross this block,  this should invoke a jump 
			//invoke a jump to top,left of visual range.
			this.loadCell(vrange.top, vrange.left, 0, null, true);
			return true;
		} else if (!(tRow >= top && lCol >= left && bRow <= bottom && rCol <= right)) {

			var fetchHeight = fetchWidth = -1,
				arRight = ar.rect.right,
				arBtm = ar.rect.bottom;
			if (arRight < rCol && arBtm < bRow) {
				//preload east and south
				fetchWidth = this._getFetchWidth(rCol);
				fetchHeight = this._getFetchHeight(bRow);
			} else if (ar.rect.right < rCol) { //preload east only
				fetchWidth = this._getFetchWidth(rCol);
				fetchHeight = ar.rect.bottom - bottom;
			} else if (ar.rect.bottom < bRow) { //preload south only
				fetchHeight = this._getFetchHeight(bRow);
				fetchWidth = ar.rect.right - right;
			}

			var token = zkS.addCallback(function () {
				//when inital , there is a loadForVisible, 
				//in this monent, use might click to invalidate this spreadhsheet
				//if click first, and then call loadForVisible.(those 2 event will send in one au_
				//then this call bak will throw exception because this block is invalidate.
				//TODO : so, in any call, i should check is it valid or not. 
				if (local.invalid) return;
				sheet.activeBlock.loadstate = zss.MainBlockCtrl.IDLE;
			});
			this._sendOnCellFetch(token, "visible", "", -1, -1, fetchWidth, fetchHeight, vrange);
			return true;
		}
		
		if (created) {
			sheet.sendSyncblock(true);
		}
		return false;
	},
	/**
	 * Reload block area content 
	 * @param string dir: "east", "west", "south", "north" or null
	 */
	reloadBlock: function (dir) {
		var local = this,
			sheet = this.sheet,
		//TODO should i control this??
		//local.loadstate!=zkSMainBlockCtrl.IDLE
			vrange = zss.SSheetCtrl._getVisibleRange(sheet),
			top = this.range.top,
			left = this.range.left,
			right = this.range.right,
			bottom = this.range.bottom; 
		var token = zkS.addCallback(function () {
			if (sheet.invald) return; 
			var block = sheet.activeBlock,
				range = block.range;
			block.loadstate = zss.MainBlockCtrl.IDLE;
			if (dir == "east" || dir == "west") {
				sheet.sp.scrollToVisible(null, dir == "east" ? range.right : range.left);
			} else if (dir == "south" || dir == "north") {
				sheet.sp.scrollToVisible(dir == "south" ? range.bottom: range.top ? range.right: range.left, null);
			}
			block.loadForVisible();
		});
		var col = dir=="east" ? right : left,
			row = dir=="south" ? bottom : top;
		this._sendOnCellFetch(token, "jump", (dir ? dir: "west"), col, row, -1, -1, vrange);
	}
}, {
	IDLE: 0,
	LOADING: 1
});
})();